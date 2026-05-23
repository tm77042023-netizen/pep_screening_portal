from __future__ import annotations

from io import BytesIO
import unittest

from openpyxl import Workbook

from app import AdverseMediaAlert, AdverseMediaSearch, CurrentAffairsIssue, DeveloperApiKey, PepRecord, PipRelationship, PublicSearchUsage, RecordAuditLog, ScreeningRequest, ScreeningResult, SubscriptionInvoice, User, app, clean_html_text_and_links, clean_profile_summary, create_relationship_records_from_text, db, extract_openai_output_text, basis_from_record_rules, extract_validated_candidates, parse_cipa_registry_text, rejection_reason_for_name, run_public_search, save_adverse_media_result, screen_name, validate_candidate


class CandidateValidationTests(unittest.TestCase):
    def assert_rejected_name(self, value: str) -> None:
        self.assertIsNotNone(rejection_reason_for_name(value), value)

    def test_rejects_known_navigation_and_institution_false_positives(self) -> None:
        rejected = [
            "Visiting Parliament Our Legislators",
            "The Speaker",
            "Technology Research Centre",
            "Speaks Petitions",
            "Latest News",
            "Bills Glossary",
            "Botswana Website Design",
            "National Assembly Quick Links",
            "Parliamentary Business Order",
            "Calendar Downloads Constitution",
            "View All Leadership Our",
            "Updates View All Upload",
            "Get Social",
            "Social Services",
            "Civil Registration",
            "Radio Services",
            "Rwanda Sign Six Bilateral",
            "Register Company",
        ]
        for value in rejected:
            with self.subTest(value=value):
                self.assert_rejected_name(value)

    def test_accepts_president_with_evidence(self) -> None:
        text = "President Duma Boko addressed Parliament on national development priorities."
        candidates, metrics = extract_validated_candidates(
            text,
            source_name="Test source",
            source_url="https://example.test",
            source_jurisdiction="Botswana",
        )
        self.assertEqual(metrics["filtered_out"], 0)
        self.assertEqual(candidates[0]["name"], "Duma Boko")
        self.assertEqual(candidates[0]["category"], "Executive government")
        self.assertGreaterEqual(candidates[0]["confidence_score"], 70)

    def test_splits_merged_names_when_role_evidence_exists(self) -> None:
        text = "President DumaBoko addressed Parliament while Vice President NdabaGaolathe attended."
        candidates, _metrics = extract_validated_candidates(
            text,
            source_name="Test source",
            source_url="https://example.test",
            source_jurisdiction="Botswana",
        )
        names = {candidate["name"] for candidate in candidates}
        self.assertIn("Duma Boko", names)
        self.assertIn("Ndaba Gaolathe", names)

    def test_accepts_honourable_former_president(self) -> None:
        text = "Hon. Seretse Khama Ian Khama, former President, attended the event."
        candidates, _metrics = extract_validated_candidates(
            text,
            source_name="Test source",
            source_url="https://example.test",
            source_jurisdiction="Botswana",
        )
        names = {candidate["name"] for candidate in candidates}
        self.assertIn("Seretse Khama Ian Khama", names)

    def test_foreign_pep_is_flagged(self) -> None:
        result = validate_candidate(
            "Viktor Orban",
            "Viktor Orban, Prime Minister of Hungary, joined the regional meeting.",
            source_name="DailyNews",
            source_url="daily.pdf",
            source_jurisdiction="Botswana",
        )
        self.assertTrue(result["accepted"])
        self.assertEqual(result["category"], "Foreign PEP/PIP")
        self.assertEqual(result["jurisdiction"], "Hungary")

    def test_known_foreign_leader_overrides_source_jurisdiction(self) -> None:
        result = validate_candidate(
            "Paul Kagame",
            "Paul Kagame joined talks during the regional summit.",
            source_name="DailyNews",
            source_url="daily.pdf",
            source_jurisdiction="Botswana",
        )
        self.assertTrue(result["accepted"])
        self.assertEqual(result["category"], "Foreign PEP/PIP")
        self.assertEqual(result["jurisdiction"], "Rwanda")

    def test_traditional_leadership_category(self) -> None:
        text = "Kgosi Boniface Obuseng addressed the kgotla."
        candidates, _metrics = extract_validated_candidates(
            text,
            source_name="Test source",
            source_url="https://example.test",
            source_jurisdiction="Botswana",
        )
        self.assertEqual(candidates[0]["name"], "Boniface Obuseng")
        self.assertEqual(candidates[0]["category"], "Traditional leadership")

    def test_full_name_without_role_is_not_saved_as_normal_candidate(self) -> None:
        result = validate_candidate(
            "Festus Mogae",
            "Festus Mogae was mentioned in passing without a public role in this sentence.",
            source_name="Test source",
            source_url="https://example.test",
            source_jurisdiction="Botswana",
        )
        self.assertFalse(result["accepted"])
        self.assertEqual(result["category"], "Public-source mention only")

    def test_basis_fallback_is_generated_for_seed_or_manual_records(self) -> None:
        record = PepRecord(
            full_name="Duma Boko",
            category="Domestic PIP",
            jurisdiction="Botswana",
            position="President",
            status="Current",
            source_name="Seed data",
        )
        self.assertIn("President", basis_from_record_rules(record))

    def test_single_surname_surfaces_unique_confirmed_record(self) -> None:
        with app.app_context():
            record, score, decision = screen_name("boko")
        self.assertIsNotNone(record)
        self.assertEqual(record.full_name, "Duma Boko")
        self.assertGreaterEqual(score, 75)
        self.assertEqual(decision, "Possible match")

    def test_profile_summary_is_kept_brief(self) -> None:
        noisy = (
            "Duma Gideon Boko is a Motswana politician and human rights lawyer serving as President of Botswana. "
            "His victory ended 60 years of Botswana Democratic Party rule. "
            "This unrelated paragraph talks about trading profits and grant calculations. "
            "More unrelated content follows after the useful profile."
        )
        summary = clean_profile_summary(noisy)
        self.assertIn("Duma Gideon Boko", summary)
        self.assertLessEqual(summary.count("."), 4)

    def test_openai_output_text_parser_supports_responses_shape(self) -> None:
        data = {
            "output": [
                {
                    "content": [
                        {
                            "type": "output_text",
                            "text": '{"candidates":[]}',
                        }
                    ]
                }
            ]
        }
        self.assertEqual(extract_openai_output_text(data), '{"candidates":[]}')

    def test_admin_page_includes_pdf_upload(self) -> None:
        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.get("/admin")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Upload PDFs", response.data)
        self.assertIn(b"Staged imports", response.data)
        self.assertIn(b"Upload and analyse PDFs", response.data)
        self.assertIn(b"Excel import / export", response.data)
        self.assertIn(b"Web link PIP/adverse-media review", response.data)
        self.assertIn(b"Review web link", response.data)

    def test_web_link_parser_extracts_text_and_same_domain_links(self) -> None:
        html = """
        <html><head><title>Leadership News</title></head>
        <body><nav>Home Links</nav><p>President Duma Boko addressed the committee.</p>
        <a href="/profile">Profile</a><a href="https://external.test/page">External</a></body></html>
        """
        title, text, links = clean_html_text_and_links(html, "https://example.test/news", same_domain_only=True, max_links=5)
        self.assertEqual(title, "Leadership News")
        self.assertIn("President Duma Boko", text)
        self.assertEqual(links, ["https://example.test/profile"])

    def test_cipa_registry_parser_extracts_company_and_directors(self) -> None:
        sample = (
            "K M B Consulting Proprietary Limited (BW00004468071)\n"
            "Company status\nRemoved\nCompany type\nPrivate Company\n"
            "Directors\n"
            "Duma Gideon Boko\nNationality\nBotswana\nAppointment Date\n04 October 2022\n"
            "Auswell Mashaba\nNationality\nSouth Africa\nAppointment Date\n13 October 2022\n"
        )
        parsed = parse_cipa_registry_text(sample)
        self.assertEqual(parsed.get("company_number"), "BW00004468071")
        directors = parsed.get("directors") or []
        self.assertEqual(len(directors), 2)

    def test_public_search_is_limited_to_three_names(self) -> None:
        with app.app_context():
            results, limited = run_public_search("Duma Boko\nNdaba Gaolathe\nFestus Mogae\nExtra Person")
        self.assertTrue(limited)
        self.assertEqual(len(results), 3)

    def test_landing_page_renders_public_search(self) -> None:
        client = app.test_client()
        response = client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Search up to 3 names", response.data)
        self.assertIn(b"Live adverse-media intelligence", response.data)
        self.assertIn(b"Screening results are decision-support", response.data)
        self.assertIn(b"not proof of wrongdoing", response.data)
        self.assertIn(b"Subscription plans", response.data)
        self.assertIn(b"Register workspace", response.data)
        self.assertNotIn(b"Demo access", response.data)
        self.assertNotIn(b"admin@example.com / admin123", response.data)
        self.assertNotIn(b"client@example.com / client123", response.data)
        response = client.post("/public-search", data={"public_names": "boko"}, follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"boko", response.data)

    def test_admin_can_create_developer_api_key(self) -> None:
        with app.app_context():
            DeveloperApiKey.query.filter_by(name="Unit test integration").delete()
            db.session.commit()
        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.post(
            "/developer-access",
            data={"action": "create_api_key", "name": "Unit test integration"},
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"One-time API key", response.data)
        self.assertIn(b"rsk_live_", response.data)
        self.assertIn(b"/api/adverse-media/search", response.data)
        with app.app_context():
            key = DeveloperApiKey.query.filter_by(name="Unit test integration").first()
            self.assertIsNotNone(key)
            db.session.delete(key)
            db.session.commit()

    def test_non_admin_cannot_open_developer_access(self) -> None:
        client = app.test_client()
        client.post("/login", data={"email": "client@example.com", "password": "client123"}, follow_redirects=True)
        response = client.get("/developer-access", follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Developer access is restricted", response.data)
        self.assertNotIn(b"Organisation API keys", response.data)

    def test_registration_creates_pending_workspace_and_invoice(self) -> None:
        email = "signup-test@example.com"
        with app.app_context():
            user = User.query.filter_by(email=email).first()
            if user:
                SubscriptionInvoice.query.filter_by(user_id=user.id).delete()
                db.session.delete(user)
                db.session.commit()
        client = app.test_client()
        response = client.post(
            "/register",
            data={
                "name": "Signup Test",
                "email": email,
                "password": "pass12345",
                "organisation": "Signup Test Org",
                "phone": "70000000",
                "plan_code": "professional",
                "billing_contact_name": "Signup Test",
                "billing_contact_email": email,
            },
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Complete subscription payment", response.data)
        self.assertIn(b"Softdayta Risk subscription invoice", response.data)
        with app.app_context():
            user = User.query.filter_by(email=email).first()
            self.assertIsNotNone(user)
            self.assertEqual(user.subscription_status, "pending")
            invoice = SubscriptionInvoice.query.filter_by(user_id=user.id).first()
            self.assertIsNotNone(invoice)
            db.session.delete(invoice)
            db.session.delete(user)
            db.session.commit()

    def test_subscription_page_renders_for_logged_in_user(self) -> None:
        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.get("/subscription")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Create Renewal Invoice", response.data)
        self.assertIn(b"Invoice History", response.data)

    def test_tender_readiness_pages_and_feed_render(self) -> None:
        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        coverage = client.get("/admin/coverage")
        self.assertEqual(coverage.status_code, 200)
        self.assertIn(b"Verified Coverage Dashboard", coverage.data)
        dictionary = client.get("/admin/data-dictionary")
        self.assertEqual(dictionary.status_code, 200)
        self.assertIn(b"Data Dictionary", dictionary.data)
        feed = client.get("/api/feed/records")
        self.assertEqual(feed.status_code, 200)
        self.assertIn(b"record_count", feed.data)
        compliance = client.get("/compliance-pack")
        self.assertEqual(compliance.status_code, 200)
        self.assertIn(b"Compliance Pack", compliance.data)

    def test_admin_can_add_relationship_and_audit_log(self) -> None:
        with app.app_context():
            principal = PepRecord.query.filter_by(full_name="Duma Boko").first()
            PipRelationship.query.filter_by(related_name="Unit Test Associate").delete()
            RecordAuditLog.query.filter_by(action="relationship_created").delete()
            db.session.commit()
            principal_id = principal.id
        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.post(
            "/admin/relationships",
            data={
                "principal_record_id": str(principal_id),
                "related_name": "Unit Test Associate",
                "relationship_type": "Close associate",
                "jurisdiction": "Botswana",
                "confidence_score": "85",
                "review_status": "Candidate review",
                "source_name": "Unit test source",
                "source_excerpt": "Relationship evidence.",
            },
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"PIP relationship / RCA record added", response.data)
        with app.app_context():
            relationship = PipRelationship.query.filter_by(related_name="Unit Test Associate").first()
            self.assertIsNotNone(relationship)
            audit = RecordAuditLog.query.filter_by(relationship_id=relationship.id, action="relationship_created").first()
            self.assertIsNotNone(audit)
            db.session.delete(audit)
            db.session.delete(relationship)
            db.session.commit()

    def test_pdf_and_web_text_can_auto_create_relationship_candidates(self) -> None:
        text = "President Duma Boko attended with his wife Jane Doe during the public event."
        with app.app_context():
            PipRelationship.query.filter_by(related_name="Jane Doe").delete()
            RecordAuditLog.query.filter_by(action="relationship_auto_extracted").delete()
            db.session.commit()
            created = create_relationship_records_from_text(
                text,
                source_name="Unit test web page",
                source_url="https://example.test/relationship",
                source_jurisdiction="Botswana",
                source_type="Web link review",
            )
            self.assertEqual(created, 1)
            relationship = PipRelationship.query.filter_by(related_name="Jane Doe").first()
            self.assertIsNotNone(relationship)
            self.assertEqual(relationship.relationship_type, "Spouse")
            self.assertEqual(relationship.principal.full_name, "Duma Boko")
            audit = RecordAuditLog.query.filter_by(relationship_id=relationship.id, action="relationship_auto_extracted").first()
            self.assertIsNotNone(audit)
            db.session.delete(audit)
            db.session.delete(relationship)
            db.session.commit()

    def test_adverse_media_result_is_saved_with_alerts(self) -> None:
        payload = {
            "searched_name": "Duma Boko",
            "jurisdiction": "Botswana",
            "overall_risk_level": "Low",
            "overall_summary": "Mentioned in official capacity; no direct allegation identified.",
            "pip_status": "Domestic PIP",
            "display_badges": ["Domestic PIP", "Official capacity only"],
            "alerts": [
                {
                    "headline": "Public-sector governance mention",
                    "risk_theme": ["Public-sector governance"],
                    "risk_level": "Low",
                    "linkage_type": "Official capacity only",
                    "summary": "The source mentions the subject in official capacity only.",
                    "source_name": "Example News",
                    "source_url": "https://example.com/news",
                    "source_date": "2026-05-16",
                    "recommended_action": "Record as contextual and continue monitoring.",
                }
            ],
        }
        with app.app_context():
            saved = save_adverse_media_result(payload, fallback_name="Duma Boko")
            self.assertEqual(saved.alerts[0].linkage_type, "Official capacity only")
            db.session.delete(saved.alerts[0])
            db.session.delete(saved)
            db.session.commit()

    def test_dashboard_renders_adverse_media_search_ui(self) -> None:
        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.get("/dashboard")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Screen a subject", response.data)
        self.assertIn(b"Adverse Media Alerts", response.data)
        self.assertIn(b"Human review is required before reliance", response.data)
        self.assertIn(b"Database record", response.data)
        self.assertNotIn(b"No matched record", response.data)
        self.assertNotIn(b"No database record linked", response.data)

    def test_dashboard_recent_screenings_only_shows_possible_matches(self) -> None:
        with app.app_context():
            user = User.query.filter_by(email="admin@example.com").first()
            record = PepRecord.query.filter_by(full_name="Duma Boko").first()
            screening = ScreeningRequest(user_id=user.id, request_type="single")
            db.session.add(screening)
            db.session.flush()
            db.session.add(
                ScreeningResult(
                    request_id=screening.id,
                    searched_name="Noise Only Person",
                    decision="No match",
                    match_score=0,
                )
            )
            db.session.add(
                ScreeningResult(
                    request_id=screening.id,
                    searched_name="Boko",
                    matched_record_id=record.id,
                    decision="Possible match",
                    match_score=75,
                )
            )
            db.session.commit()
            screening_id = screening.id
        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.get("/dashboard")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Recent Possible Matches", response.data)
        self.assertIn(b"Boko", response.data)
        self.assertNotIn(b"Noise Only Person", response.data)
        with app.app_context():
            screening = db.session.get(ScreeningRequest, screening_id)
            for result in list(screening.results):
                db.session.delete(result)
            db.session.delete(screening)
            db.session.commit()

    def test_public_search_limit_is_enforced_by_ip(self) -> None:
        with app.app_context():
            PublicSearchUsage.query.delete()
            db.session.commit()
        client = app.test_client()
        first = client.post(
            "/public-search",
            data={"public_names": "Duma Boko\nNdaba Gaolathe\nFestus Mogae"},
            environ_base={"REMOTE_ADDR": "198.51.100.10"},
        )
        self.assertEqual(first.status_code, 200)
        self.assertIn(b"Remaining public preview searches today", first.data)
        self.assertIn(b"<strong>0</strong>", first.data)
        self.assertIn(b"Preview limit reached", first.data)
        second = client.post(
            "/public-search",
            data={"public_names": "Extra Person"},
            environ_base={"REMOTE_ADDR": "198.51.100.10"},
            follow_redirects=True,
        )
        self.assertEqual(second.status_code, 200)
        self.assertIn(b"Your public preview limit has been used for today", second.data)

    def test_admin_can_add_current_affairs_issue(self) -> None:
        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.post(
            "/admin/current-affairs",
            data={
                "title": "Unit test politics issue",
                "category": "Politics",
                "jurisdiction": "Botswana",
                "summary": "A test issue for public-office monitoring.",
                "is_active": "1",
            },
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        with app.app_context():
            issue = CurrentAffairsIssue.query.filter_by(title="Unit test politics issue").first()
            self.assertIsNotNone(issue)
            db.session.delete(issue)
            db.session.commit()

    def test_admin_can_export_records_workbook(self) -> None:
        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.get("/admin/records/export")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"PK", response.data[:4])

    def test_admin_can_import_excel_record(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "PEP Records"
        sheet.append([
            "id",
            "full_name",
            "category",
            "jurisdiction",
            "position",
            "status",
            "adverse_media_status",
            "adverse_media_linkage",
            "source_type",
            "source_name",
            "source_excerpt",
            "reviewer_notes",
            "notes",
        ])
        sheet.append([
            "",
            "Excel Import Test Person",
            "Domestic PIP",
            "Botswana",
            "Councillor",
            "Candidate review",
            "Adverse media found",
            "Contextual",
            "News",
            "Unit test workbook",
            "Evidence excerpt from workbook.",
            "Needs review by analyst.",
            "Imported for testing.",
        ])
        output = BytesIO()
        workbook.save(output)
        output.seek(0)

        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.post(
            "/admin/records/import",
            data={"records_file": (output, "records.xlsx")},
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        with app.app_context():
            record = PepRecord.query.filter_by(full_name="Excel Import Test Person").first()
            self.assertIsNotNone(record)
            self.assertEqual(record.adverse_media_status, "Adverse media found")
            db.session.delete(record)
            db.session.commit()

    def test_admin_can_confirm_candidate_record(self) -> None:
        with app.app_context():
            record = PepRecord(
                full_name="Test Confirm Person",
                category="Domestic PIP",
                jurisdiction="Botswana",
                position="Minister",
                status="Candidate review",
                source_name="Unit test",
            )
            db.session.add(record)
            db.session.commit()
            record_id = record.id
        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.post(f"/admin/records/{record_id}/action", data={"action": "confirm", "reviewer_note": "Evidence reviewed."}, follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        with app.app_context():
            updated = db.session.get(PepRecord, record_id)
            self.assertEqual(updated.status, "Confirmed")
            self.assertEqual(updated.adverse_media_status, "No adverse media")
            db.session.delete(updated)
            db.session.commit()

    def test_admin_can_update_adverse_media(self) -> None:
        with app.app_context():
            record = PepRecord(
                full_name="Test Adverse Person",
                category="Domestic PIP",
                jurisdiction="Botswana",
                position="MP",
                status="Confirmed",
                source_name="Unit test",
            )
            db.session.add(record)
            db.session.commit()
            record_id = record.id
        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.post(
            f"/admin/records/{record_id}/action",
            data={
                "action": "adverse_media",
                "adverse_media_status": "Under investigation",
                "adverse_media_linkage": "Contextual",
                "reviewer_note": "Article requires review.",
            },
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        with app.app_context():
            updated = db.session.get(PepRecord, record_id)
            self.assertEqual(updated.adverse_media_status, "Under investigation")
            self.assertEqual(updated.adverse_media_linkage, "Contextual")
            db.session.delete(updated)
            db.session.commit()

    def test_admin_can_edit_suggested_record_info(self) -> None:
        with app.app_context():
            record = PepRecord(
                full_name="Suggested Info Test",
                category="Domestic PIP",
                jurisdiction="Botswana",
                position="Unknown",
                status="Candidate review",
                source_name="Unit test",
            )
            db.session.add(record)
            db.session.commit()
            record_id = record.id
        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.post(
            f"/admin/records/{record_id}/edit",
            data={
                "full_name": "Suggested Info Test Updated",
                "category": "Foreign PEP/PIP",
                "jurisdiction": "Rwanda",
                "position": "President",
                "status": "Needs review",
                "source_excerpt": "Corrected source evidence.",
                "adverse_media_status": "Under investigation",
                "reviewer_notes": "Corrected by analyst.",
                "notes": "Manual correction.",
            },
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        with app.app_context():
            updated = db.session.get(PepRecord, record_id)
            self.assertEqual(updated.full_name, "Suggested Info Test Updated")
            self.assertEqual(updated.category, "Foreign PEP/PIP")
            self.assertEqual(updated.jurisdiction, "Rwanda")
            self.assertEqual(updated.position, "President")
            self.assertEqual(updated.status, "Needs review")
            self.assertEqual(updated.source_excerpt, "Corrected source evidence.")
            self.assertEqual(updated.adverse_media_status, "Under investigation")
            db.session.delete(updated)
            db.session.commit()

    def test_admin_can_manually_activate_subscription(self) -> None:
        with app.app_context():
            user = User(
                name="Pending Subscriber",
                email="pending_subscriber@example.com",
                organisation="Unit test org",
                role="subscriber",
                subscription_status="pending",
                plan_code="professional",
            )
            user.set_password("test12345")
            db.session.add(user)
            db.session.commit()
            user_id = user.id

        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.post(
            "/developer-access",
            data={"action": "activate_subscription", "user_id": str(user_id), "plan_code": "professional", "clear_trial": "1"},
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        with app.app_context():
            updated_user = db.session.get(User, user_id)
            self.assertIsNotNone(updated_user)
            self.assertEqual(updated_user.subscription_status, "active")
            db.session.delete(updated_user)
            db.session.commit()

    def test_admin_results_collapse_core_record_fields(self) -> None:
        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.get("/admin")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b'class="field-collapse"', response.data)
        self.assertIn(b'class="collapsed-fields"', response.data)
        self.assertIn(b"&gt;", response.data)

    def test_admin_rejected_tab_displays_rejected_review_records(self) -> None:
        with app.app_context():
            record = PepRecord(
                full_name="Rejected Review Test Person",
                category="Domestic PIP",
                jurisdiction="Botswana",
                position="Councillor",
                status="Rejected / not a person",
                source_name="Unit test",
            )
            db.session.add(record)
            db.session.commit()
            record_id = record.id
        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.get("/admin?tab=rejected")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Rejected Review Test Person", response.data)
        with app.app_context():
            record = db.session.get(PepRecord, record_id)
            RecordAuditLog.query.filter_by(record_id=record_id).delete()
            db.session.delete(record)
            db.session.commit()

    def test_admin_bulk_edit_restores_rejected_record_and_confirms_relationship(self) -> None:
        with app.app_context():
            principal = PepRecord(
                full_name="Bulk Principal PIP",
                category="Domestic PIP",
                jurisdiction="Botswana",
                position="Minister",
                status="Current",
                source_name="Unit test",
            )
            rejected = PepRecord(
                full_name="Bulk Rejected Candidate",
                category="Public-source mention only",
                jurisdiction="Botswana",
                position="",
                status="Rejected / not a person",
                source_name="Unit test",
            )
            db.session.add_all([principal, rejected])
            db.session.flush()
            relationship = PipRelationship(
                principal_record_id=principal.id,
                related_name="Bulk Related Person",
                relationship_type="Close associate",
                category="Related party",
                jurisdiction="Botswana",
                source_name="Unit test",
                confidence_score=45,
                review_status="Candidate review",
            )
            db.session.add(relationship)
            db.session.commit()
            principal_id = principal.id
            rejected_id = rejected.id
            relationship_id = relationship.id

        client = app.test_client()
        client.post("/login", data={"email": "admin@example.com", "password": "admin123"}, follow_redirects=True)
        response = client.post(
            "/admin/bulk-edit",
            data={
                "record_id": [str(rejected_id)],
                f"record_{rejected_id}_full_name": "Bulk Restored Candidate",
                f"record_{rejected_id}_category": "Domestic PIP",
                f"record_{rejected_id}_position": "Councillor",
                f"record_{rejected_id}_jurisdiction": "Botswana",
                f"record_{rejected_id}_status": "Confirmed",
                f"record_{rejected_id}_verification_status": "Source verified",
                f"record_{rejected_id}_source_reliability": "High",
                f"record_{rejected_id}_adverse_media_status": "No adverse media",
                f"record_{rejected_id}_source_name": "Council minutes",
                f"record_{rejected_id}_source_url": "https://example.test/council",
                f"record_{rejected_id}_source_excerpt": "Named as councillor in public minutes.",
                f"record_{rejected_id}_last_verified_date": "2026-05-19",
                f"record_{rejected_id}_next_review_due": "2026-11-19",
                f"record_{rejected_id}_reviewer_notes": "Approved after evidence review.",
                "relationship_id": [str(relationship_id)],
                f"relationship_{relationship_id}_related_name": "Bulk Related Person",
                f"relationship_{relationship_id}_relationship_type": "Close associate",
                f"relationship_{relationship_id}_category": "Related party",
                f"relationship_{relationship_id}_jurisdiction": "Botswana",
                f"relationship_{relationship_id}_confidence_score": "82",
                f"relationship_{relationship_id}_review_status": "Confirmed",
                f"relationship_{relationship_id}_source_name": "Public declaration",
                f"relationship_{relationship_id}_source_url": "https://example.test/declaration",
                f"relationship_{relationship_id}_source_excerpt": "Relationship disclosed in public document.",
                f"relationship_{relationship_id}_reviewer_notes": "Confirmed by admin.",
            },
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        with app.app_context():
            updated = db.session.get(PepRecord, rejected_id)
            updated_relationship = db.session.get(PipRelationship, relationship_id)
            self.assertEqual(updated.full_name, "Bulk Restored Candidate")
            self.assertEqual(updated.status, "Confirmed")
            self.assertEqual(updated.verification_status, "Source verified")
            self.assertEqual(updated.source_url, "https://example.test/council")
            self.assertEqual(updated_relationship.review_status, "Confirmed")
            self.assertEqual(updated_relationship.confidence_score, 82)
            self.assertEqual(updated_relationship.source_url, "https://example.test/declaration")
            RecordAuditLog.query.filter(RecordAuditLog.record_id.in_([principal_id, rejected_id])).delete(synchronize_session=False)
            RecordAuditLog.query.filter_by(relationship_id=relationship_id).delete()
            db.session.delete(updated_relationship)
            db.session.delete(updated)
            db.session.delete(db.session.get(PepRecord, principal_id))
            db.session.commit()


if __name__ == "__main__":
    unittest.main()
