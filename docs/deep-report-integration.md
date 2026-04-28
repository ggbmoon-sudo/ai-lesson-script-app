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
| Local citation index | Published revisions build chunk IDs, source hashes, token vectors, confidence scores |
| Role modes | Teacher / TA / Student / Admin front-end permission simulation |
| QA metrics | Grounded rate, refusal rate, helpful feedback, and needs-teacher counts |

## Still Missing

| Gap | Why It Matters |
|---|---|
| Real vector RAG | Current retrieval is local lexical + cosine vector approximation, not embedding search |
| OCR and layout extraction | Scanned PDFs and complex tables need OCR/layout models |
| RBAC roles | Teacher, TA, student, admin permissions are front-end simulated only |
| Personal cloud sync | Google Drive backup supports manual restore and optional auto-backup, but not conflict-aware sync |
| LMS integration | No Canvas / Google Classroom / LTI workflow yet |
| Speech path | No STT/TTS or realtime voice assistant yet |
| Evaluation suite | Basic QA metrics exist, but no teacher gold set, red-team set, or latency tracking |

## Suggested Next Build Order

1. Replace local citation vectors with embeddings and persistent vector storage.
2. Move role permissions to a backend policy layer.
3. Add automatic cloud backup scheduling and conflict detection for Google Drive.
4. Add OCR provider abstraction for future Azure / Google / AWS integration.
5. Add LMS export or import stubs.
6. Add teacher gold-set and red-team evaluation runners.
