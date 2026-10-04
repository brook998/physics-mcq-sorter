# Changelog

## v3
- Filter Hindi/Devanagari at the full question-region level.
- Require recognizable MCQ options before accepting a numbered question.
- Remove visible source metadata from DOCX output.
- Normalize question numbering sequentially within each chapter.
- Detect paper set codes such as 55/4/1 and print them at the bottom-right of each MCQ image.
- Improve chapter classification by preventing short keyword substring false positives.

## v2
- Support up to 21 PDF papers per run.
- Faster native-text-first processing with OCR fallback.
