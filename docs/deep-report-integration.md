# Deep Research Report Integration

This note maps `deep-research-report.md` into the current EduScript AI Studio implementation.

## Product Shift

The new report reframes the app as a teaching material operating system rather than a PDF chat tool. The product should transform source files into a traceable intermediate layer, then generate editable materials, timed lecture scripts, and grounded assistants.

## Implemented From The Report

| Report Requirement | Current Implementation |
|---|---|
| Web-first teacher workflow | Static web app plus optional Node server |
| Separate import, planning, editing, interaction layers | Workspace sections for builder, script, assistant, student QA, library |
| PPTX/DOCX/PDF parsing | Server-side basic parser for text-bearing files |
| XLSX / schedule import | Server-side worksheet text extraction |
| PPTX renderer from structured slides | `/api/export-pptx` creates editable PowerPoint package |
| AI generates draft, teacher publishes | `publishLesson()` creates a published revision snapshot |
| Student assistant only sees published content | Student QA reads `publishedRevision` only |
| Grounded answers with source display | Student QA returns slide/material sources or refuses |
| Provenance and audit | Audit log records generation, parsing, publishing, export, feedback |
| Source refs | Published slides are annotated with best-effort material refs |

## Still Missing

| Gap | Why It Matters |
|---|---|
| Real vector RAG | Current retrieval is keyword scoring, not embedding search |
| OCR and layout extraction | Scanned PDFs and complex tables need OCR/layout models |
| RBAC roles | Teacher, TA, student, admin permissions are still simulated |
| LMS integration | No Canvas / Google Classroom / LTI workflow yet |
| Speech path | No STT/TTS or realtime voice assistant yet |
| Evaluation suite | No teacher gold set, red-team set, groundedness metrics, or latency tracking |

## Suggested Next Build Order

1. Add local vector-like index with chunk IDs, source hashes, and citation precision fields.
2. Add role mode switch: Teacher / Student / TA / Admin.
3. Add assessment and QA metrics dashboard.
4. Add OCR provider abstraction for future Azure / Google / AWS integration.
5. Add LMS export or import stubs.
