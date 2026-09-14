# 📄 Document README — Architecture & Build Guide

**File:** `GuardrailIQ_Architecture_and_Build_Guide.docx`

This is the formal design document for the GuardrailIQ solution. It is a
7-page, professionally formatted Word document with a cover page, table of
contents, page numbering, running header, diagrams and tables.

## Contents at a glance

| # | Section | What it covers |
|---|---------|----------------|
| 1 | **Executive Overview** | What the tool does and why it is deterministic / LLM-free |
| 2 | **Solution Architecture** | Pipeline diagram + per-stage responsibility table |
| 3 | **Core Data Model** | The five dataclasses that keep citations intact end-to-end |
| 4 | **Key Algorithms** | Severity/concept tagging, TF-IDF+cosine+NetworkX de-dup, conflict detection, negation-aware evaluation |
| 5 | **Setup & Run Steps** | Prerequisites, commands, first-run workflow |
| 6 | **Application Walkthrough** | What each of the 7 UI tabs shows |
| 7 | **Design Decisions & Governance** | Traceability, human-in-the-loop, no-fabrication guarantees |
| 8 | **Mapping to Challenge Requirements** | Each mandatory outcome → where it is met |
| 9 | **Validation Results** | The 12/12 design-classification table + seeded scenarios reproduced |
| 10 | **Assumptions & Limitations** | Scope boundaries and extension points |

## How to read / edit

- Open in **Microsoft Word** (or any `.docx`-compatible viewer).
- The **Table of Contents** is clickable; right-click → *Update Field* to
  refresh page numbers after editing.
- The architecture diagram is also available as a standalone image at
  `../diagram/architecture_diagram.png`.

## Related files

- Application source: `../app/guardrailiq_app.py`
- Project overview & run instructions: `../README.md`
