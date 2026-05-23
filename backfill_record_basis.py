from __future__ import annotations

from app import PepRecord, app, append_basis_if_missing, basis_from_record_rules, db


def main() -> None:
    with app.app_context():
        updated = 0
        records = PepRecord.query.filter(
            PepRecord.status != "Rejected / not a person",
            PepRecord.category != "Rejected / not a person",
        ).all()
        for record in records:
            original_notes = record.notes or ""
            record.notes = append_basis_if_missing(original_notes, basis_from_record_rules(record))
            if record.notes != original_notes:
                updated += 1
        db.session.commit()
        print(f"{updated} non-rejected record(s) updated with basis notes.")


if __name__ == "__main__":
    main()
