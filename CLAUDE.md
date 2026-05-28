# Rimon Health Internship — AI Solutions

> **FOR CLAUDE:** Read this file at the start of every session. This is the running context for Komala's internship at Rimon Health. Pick up exactly where we left off. Use caveman ultra mode throughout.

**Start:** ~2026-04-21 | **Role:** AI integration + workflow automation
**Compliance:** HIPAA · BAA · PHI mandatory on all solutions

---

## HIPAA/PHI Ground Rules

| Rule | Detail |
|------|--------|
| No raw PHI → non-BAA LLMs | De-identify first OR use BAA-covered endpoint |
| BAA providers | Azure OpenAI ✓, AWS Bedrock ✓, Google Vertex AI ✓, OpenAI Enterprise ✓ |
| Storage | Encrypted at rest + transit, RBAC, audit logs |
| Testing | Use synthetic/de-identified data only during dev |
| Anthropic Claude API | BAA available via Enterprise plan |

---

## Rimon Health — What They Actually Do (From Research)

**Services:** ASD evaluations, Learning Disabilities (Dyslexia/Dyscalculia/Dysgraphia/ADHD), Cognitive & Memory (Alzheimer's/Dementia/TBI), Forensic Psychology
**Report framework:** Brooke's Comments — Action Checklist (sections: History, Test Results, Diagnosis, Recommendations)
**Key insight:** Reports need to be modular per condition type — not one-size-fits-all

---

## Prototype Design — Automated Report Writer

### Modules (one per eval type):
| Module | Conditions | Key Inputs |
|--------|-----------|------------|
| Neurodevelopmental | ASD | Behavioral obs, social interaction, parent/teacher questionnaire scores |
| Learning Disability | Dyslexia, Dyscalculia, Dysgraphia, ADHD | Reading/math/processing speed standardized scores |
| Cognitive & Memory | Alzheimer's, Dementia, TBI | Memory recall, cognitive decline metrics |
| Forensic | Forensic psych | Interview data, testimony, behavioral observations |

### Features to build (in order):
1. **Module selector** → clinician picks eval type first
2. **Score input form** → raw scores in → auto-convert to scaled + percentile + interpretation
3. **Checklist validator** → all sections filled before report generates (Brooke's framework)
4. **Report generator** → LLM drafts narrative per section
5. **Recommendation engine** → based on severity → suggest accommodations (extra time, quiet env, etc.)
6. **PHI de-id layer** → name/DOB replaced with placeholders before API call → re-injected after

### Report sections (from Brooke's Action Checklist):
- Reason for Referral
- Background & Social History (pulls from parent/teacher input)
- Behavioral Observations
- Test Results (auto-interpreted scores)
- Diagnosis / Impressions
- Recommendations + Accommodations

### HIPAA safeguards in prototype:
- E2E encryption (HTTPS + encrypted storage)
- MFA + RBAC login
- Auto session timeout
- Audit logs (who accessed, when, what changed)
- BAA-covered API only (Azure OpenAI or Anthropic Enterprise)
- PHI stitched back only at final render, inside secure env
- Consent forms attached before report marked "Done"
- Temp processing files wiped after generation

### Special: Forensic module extras
- Contradiction detection → flag inconsistencies between interview + test data
- Professional tone enforcement → NLP for "absolute truth" forensic language

---

## ✅ CONFIRMED PROBLEMS TO SOLVE (Manager-verified)

### Problem 1: Automated Neuropsych Report Writing
- **Status:** ✅ Prototype DONE
- **Location:** `Desktop/rimon-prototype/app.py`
- **Stack:** Python + Streamlit + Gemini API (demo) → swap to Azure OpenAI with BAA for production
- **Modules built:** Learning Disability, ASD, Cognitive & Memory
- **Features:** Score input → auto-interpretation → checklist validator → AI report → download
- **PHI approach:** Patient ID only (no real names) → checklist enforces de-id before generation
- **Cost:** ~$0.03/report → ~$2/month at 10 reports/week on Azure OpenAI
- **Next step:** Show manager → get feedback → refine modules based on their actual batteries used
- **API used (demo):** Groq (free) with Llama-3.3-70b-versatile
- **API for production:** Azure OpenAI GPT-4o with BAA
- **Test run:** LD module confirmed working. Report quality: clinical-grade draft, accurate interpretations, correct flagging.

### Problem 2: Appointment Scheduling + Reminders
- **Status:** Confirmed pain point — manager mentioned in interview
- **Plan:**
  1. Day 1 → check if their EHR (SimplePractice/TherapyNotes) already has reminders → just configure it
  2. If no EHR → Zapier (no-code) + Google Calendar + Twilio SMS (~$20/month, HIPAA tier)
  3. If custom needed → Python + Twilio with BAA
- **Next step:** Ask Day 1 questions first before building anything

---

## All Solutions to Build (Priority Order)

### 🥇 1. Automated Neuropsych Report Writer
**Impact: highest. Most time-consuming task for clinicians.**

- **Input:** structured score sheet (JSON/form) — test name, subtest, raw score, scaled score, percentile, domain
- **Process:** LLM (Azure OpenAI gpt-4o or Claude via Enterprise) + prompt template per report section
- **Output:** draft narrative → clinician edits + signs off
- **Sections to automate:** Reason for Referral, Background, Behavioral Observations, Test Results, Impressions, Recommendations
- **Stack:** Python + Streamlit UI, Azure OpenAI (BAA), Jinja2 templates for structure
- **PHI handling:** patient name/DOB as placeholders injected post-generation, or full BAA endpoint

---

### 🥈 2. Structured Score Input Tool
**Replaces manual score entry. Feeds into Report Writer.**

- Web form (Streamlit or simple Next.js) where clinician enters:
  - Assessment battery (WAIS-IV, WISC-V, MoCA, NEPSY-II, D-KEFS, Conners-4, MMSE, BRIEF, etc.)
  - Per-subtest: raw → scaled → percentile (auto-lookup from normative tables)
- Output: structured JSON → feeds Report Writer
- Bonus: flag scores below clinical thresholds automatically (e.g., scaled score ≤ 7 = below average)

---

### 🥉 3. Intake Form → Structured Data Pipeline
**Eliminates manual re-entry from paper/PDF intakes.**

- OCR + LLM extraction: intake PDF → structured JSON (demographics, chief complaint, history, medications, prior evals)
- Tool: **Azure Document Intelligence** (HIPAA-covered) for OCR, then GPT-4o for structuring
- Output: pre-populated patient record / report background section
- De-id layer if pushing to non-BAA endpoint

---

### 4. Normative Data RAG Assistant
**Clinician reference tool — "what does this score mean?"**

- Ingest test manuals + normative tables into vector DB
- Query: "WAIS-IV Processing Speed Index = 78, 7th percentile — clinical meaning?"
- Stack: LlamaIndex + Azure AI Search (HIPAA-covered) + Azure OpenAI
- No PHI in this tool — purely reference data

---

### 5. PHI De-identification Utility
**Enables safe use of any LLM without full BAA setup.**

- NER → detect + redact: names, DOB, SSN, address, MRN, phone, dates
- Tools: **AWS Comprehend Medical** (HIPAA) or spaCy + custom clinical NER
- Replace with placeholders: `[PATIENT_NAME]`, `[DOB]`, etc.
- Re-inject after LLM response

---

### 6. Clinical Note Summarizer
**Pre-session prep — summarize prior notes in 30 sec.**

- Input: prior session notes (text)
- Output: bullet summary — presenting concerns, test history, key findings, open recommendations
- Must run on BAA endpoint
- Simple Streamlit tool

---

### 7. Scheduling + Reminder Automation (Low-code)
**Reduce no-shows, admin overhead.**

- Automate appointment reminders via HIPAA-compliant messaging
- Integrate with existing EHR if possible (SimplePractice API, TherapyNotes webhooks)
- Tools: Zapier HIPAA tier OR custom Python + Twilio (with BAA)

---

## Tech Stack

| Layer | Tool | HIPAA |
|-------|------|-------|
| LLM | Azure OpenAI gpt-4o | ✓ BAA |
| LLM alt | Anthropic Claude Enterprise | ✓ BAA |
| OCR/docs | Azure Document Intelligence | ✓ |
| NER/de-id | AWS Comprehend Medical | ✓ |
| Vector DB | Azure AI Search | ✓ |
| Backend | Python + FastAPI | self-managed |
| UI | Streamlit (fast) or Next.js | self-managed |
| Storage | Azure Blob / AWS S3 encrypted | ✓ with config |

---

## Neuropsych Assessment Batteries IN APP

| Battery | Age Range | Mode | Status |
|---------|-----------|------|--------|
| WPPSI-IV | 2y6m–7y7m | Standard + Non-verbal | ✅ Live |
| WISC-V | 6–16yr | Standard + Non-verbal | ✅ Live |
| WAIS-V | 16–90yr | Standard + Non-verbal | ✅ Live |
| ADOS-2 | Any | All modules | ✅ Live |
| BASC-3 | Child/Adol/Preschool | PRS | ✅ Live (Q-Global DOCX parse) |
| Vineland-3 | Any | Caregiver form | ✅ Live (Q-Global DOCX parse) |
| C-TONI-2 | 6–89yr | Non-verbal IQ | ✅ Live |
| P-TONI | 3–9yr | Non-verbal IQ | ✅ Live |
| Stanford-Binet | — | — | ⏳ Deferred |

---

## Day 1 Questions for Rimon Health

- [ ] EHR/EMR in use? (SimplePractice / TherapyNotes / Epic / paper?)
- [ ] Most common assessment batteries?
- [ ] Current report writing process? (dictation / manual / template?)
- [ ] Where is patient data stored? (cloud / on-prem?)
- [ ] Existing BAA with any cloud vendor?
- [ ] Biggest clinician time sink today?
- [ ] Any existing automation (even Excel macros)?
- [ ] IT/security contact?

---

## Session Log

| Date | Work |
|------|------|
| 2026-04-16 | Initial solution space mapped. CLAUDE.md created. 7 solutions prioritized. |
| 2026-04-16 | Narrowed to 2 confirmed problems: report writing + scheduling. Prototype plan set. Building UI tomorrow. |
| 2026-04-17 | Deep research done on Rimon Health services + Brooke's Action Checklist. 4 eval modules defined. Full feature set + HIPAA safeguards documented. Prototype scope locked. |
| 2026-04-17 | Prototype BUILT. Streamlit app live at Desktop/rimon-prototype/app.py. Uses Gemini API (demo). 3 modules: LD, ASD, Cognitive. Checklist validator, score interpreter, report generator, download button. Cost breakdown done: ~$0.03/report → ~$2/month on Azure OpenAI at 10 reports/week. Manager pitch prepared. |
| 2026-04-18 | Prototype WORKING end-to-end. Switched from Gemini (quota issues) → Groq (free, fast). Tested with LD module — 17yo patient, full 6-section report generated in ~10 seconds. Score interpretation accurate (FSIQ 66 → Extremely Low, WMI/PSI flagged as clinically significant, Conners normal → no ADHD). Recommendations matched weak areas. Ready to show manager. |
| 2026-04-19 | Added PDF export (fpdf2), email (Gmail SMTP + App Password), print (opens Preview → system dialog). Fixed PDF two-column layout bug (switched multi_cell → write()). Fixed unicode crash (latin-1 encode fallback). Email personalized — clinician name, recipient name, custom note, professional sign-off. All 3 export options working. |
| 2026-05-10 | Added PDF + DOCX upload support. Added in-app audio recording. Added Q-Global DOCX parser for BASC-3 + Vineland-3 (deterministic, no AI). Added score tables to generated DOCX. Improved clinical language to match Rimon template. |
| 2026-05-12 | Fixed ADOS-2 Comparison Score crash (int(None) TypeError). Fixed ordinal suffixes (1th→1st). Fixed BASC composite CI (was ±5 estimate, now reads actual CI from parse). Fixed respondent blank fallback. Fixed Vineland grammar ("a within"). Fixed BASC form type detection for Adolescent. Fixed rec_medical_followup gating. Full sanity check + 2 test report review cycles. |
| 2026-05-19 | Added clinician dropdown (examiner + supervisor with NPI/license auto-fill). Added food preferences field. Added Q-Global smart extraction for BASC-3 + Vineland-3. Added genetics referral recommendation. Added IEP advocacy recommendation. |
| 2026-05-25 | Added WAIS-V (non-verbal mode, adult 16-90yr). Updated battery to Fifth Edition (VCI/VSI/FRI/WMI/PSI). Added C-TONI-2 block. Added P-TONI block. Updated WAIS narrative to match clinical template exactly (adult language, vocational framing, concluding paragraph). Updated CLAUDE.md. |
| 2026-05-25 | Clinician feedback fixes: (1) Auto-save session to .rimon_autosave.json on every rerun; restore banner on reload. (2) Multi-clip audio recording — each clip appends to segment list, combined transcript feeds extraction. (3) Auto-switch to Manual Entry tab after upload/extraction via tab reorder + st.rerun(). (4) Removed all em-dashes from app and report output — strip_emdashes() now handles spaced/bare em+en-dash, HTML entities; applied to all AI outputs. |
| 2026-05-28 | Clinician feedback: (1) Per-clinician indefinite save — login now collects username; autosave keyed to .rimon_save_{username}.json; no expiry. (2) Recording segments persist across sessions via autosave (already JSON-serializable). (3) Confirmed battery split: BASC-3 + Vineland-3 keep Q-Global DOCX; WPPSI-IV/WISC-V/WAIS-V/C-TONI-2/P-TONI use image/PDF upload; ADOS-2 pending Dr. Kirby confirmation. (4) Improved cognitive extraction prompts in app_preview.py — WPPSI/WISC/WAIS-specific prompts return standardized keys (VCI/VSI/FRI/WMI/PSI/FSIQ); filename-based battery detection; null filtering. |

---

## Key Architecture Notes

- **app.py** — single-file Streamlit app, ~3500+ lines
- **Non-verbal mode** — toggled per cognitive battery; replaces index scores with subtest-level inputs; generates estimated composite table + 4 narrative paragraphs
- **Q-Global parsing** — `parse_basc_docx()` and `parse_vineland_docx()` — deterministic XML/table parse, zero AI
- **OCR fallback** — Groq vision `meta-llama/llama-4-scout-17b-16e-instruct` for image score sheets
- **DOCX generation** — python-docx; `_shd_cell`, `_add_body`, `_add_subheading` helpers; all tables use Times New Roman 8pt
- **Key helpers** — `ordinal(n)`, `ss_to_pct()`, `ss_to_label()`, `_ados_int()`, `strip_emdashes()`
- **Session state guard** — `use_wais = use_wais if 'use_wais' in dir() else False` prevents NameError on rerun before checkbox renders
- **Auto-save** — `_autosave()` runs top of every rerun; save file keyed per clinician: `.rimon_save_{username}.json`; no expiry — data kept indefinitely; restore banner on fresh login
- **Tab auto-switch** — tabs reorder on extract: `if _data: manual, upload = st.tabs(...)` so Manual Entry is active tab after extraction
- **Multi-clip audio** — `recording_segments` list in session state; each clip appended, combined transcript joins all segments; per-clip delete + clear all

---

## Pending / Deferred

| Item | Status |
|------|--------|
| Stanford-Binet | Deferred - user doesn't have template yet |
| Google Drive integration | Waiting on Rimon Google account access |
| Template-matched report format | Planned |
