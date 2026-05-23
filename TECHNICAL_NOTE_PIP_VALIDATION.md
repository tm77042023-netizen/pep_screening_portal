# PIP/PEP Candidate Validation Update

## Summary

The ingestion pipeline now prioritises precision before adding scraped records to candidate review. Public websites and DailyNews PDFs pass extracted strings through a shared validation layer that checks whether the value is likely to be a human name, whether public-office context exists nearby, and whether the candidate should be treated as domestic, foreign, international/public-source mention, or rejected.

## Validation Rules Added

- Rejects website navigation, page headings, menu labels, generic parliamentary terms, and institution-only names.
- Requires person-name structure with two to four name-like capitalised tokens.
- Rejects strings with too many words, uppercase section headers, unusually long concatenated tokens, and generic non-name words.
- Extracts nearby role/title evidence and maps it to PIP categories.
- Adds confidence scoring and only saves candidates above the configured threshold.
- Logs rejected and low-confidence values separately in `RejectedCandidate`.
- Keeps scraped names in `Candidate review`; no scraped record is silently promoted to confirmed/current status.

## Ingestion Metrics

DailyNews imports now include raw strings, filtered-out values, confidence bands, and foreign PEP/PIP counts in the import log message. Public-source updates now report candidates created, rejected values, and foreign PEP/PIP flags.

## Cleanup

Use `python cleanup_false_positives.py` to mark existing likely non-person candidate-review records as `Rejected / not a person`.

Use `python cleanup_false_positives.py --delete` to delete those suspect candidate records. Seed data and confirmed records are excluded from cleanup.

## Remaining Risks

- Name extraction remains heuristic and may miss uncommon naming formats.
- Some parliamentary text may require source-specific parsing if PDFs contain broken OCR or merged words.
- Jurisdiction detection is based on nearby country mentions and should be reviewed by a compliance officer before reliance.
