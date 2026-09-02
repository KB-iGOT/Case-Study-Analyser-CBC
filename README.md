# CBC-India AGK Case Study Suite

An AI-powered Streamlit application built as a **functional prototype** for the CBC-India / AGK case-study ecosystem. The tool validates the end-to-end user workflows, prompt design, and AI feature set that are intended to be integrated into the AGK Platform. The front end is intentionally lightweight — the reusable value lies in the **prompts, user flows, and AI pipeline logic**.

## The Three Tools

| Tool | Purpose |
|---|---|
| **Case Study Analyser** | Evaluates uploaded case studies (PDF/DOCX) against the CBC-India AGK Case Study Review Rubric — weighted scoring across four assessment areas, detailed AI feedback, tiered grading, PDF report generation, and assessment history. |
| **CaseConnect** | AI-enabled case discovery: recommends AGK repository case studies to faculty based on course outlines and criteria (learners, objectives, competencies, duration, sector), with discussion points, key themes, and iGOT platform links. |
| **Case Study Generator** | Drafts new case studies from raw source material (transcripts, reports, URLs, notes) via an 8-step wizard. Supports Lesson-Drawing, Decision-Forcing, and Caselet formats, with section-by-section AI drafting, narrative continuity between sections, compliance review, teaching-note generation, and DOCX/PDF/TXT export. |

## Tech Stack at a Glance

- **Frontend + Backend**: Python / Streamlit (single-page app, session-state driven)
- **AI Engine**: Pluggable — OpenAI API (`gpt-4o`, default) or **Google Gemini** (Vertex AI / Gemini Developer API), selected at runtime with the `USE_GEMINI` flag. Prompt-based, no fine-tuning. See [Choosing the AI Provider](#choosing-the-ai-provider) below and [docs/DEVELOPER.md](docs/DEVELOPER.md#ai-engine--model-architecture) for full details.
- **Database**: PostgreSQL via SQLAlchemy (SQLite fallback for local testing)
- **Exports**: `python-docx` (DOCX), `fpdf` (PDF)
- **Auth**: Username/password with bcrypt_sha256 hashing

## Quick Start

```bash
# 1. Install dependencies (Python 3.11+)
pip install -e .          # or: uv sync

# 2. Set required environment variables
export OPENAI_API_KEY=sk-...
export DATABASE_URL=postgresql://...   # omit to fall back to local SQLite

# 3. Run
streamlit run app.py --server.port 5000
```

To run on Gemini instead of OpenAI, replace step 2 with the Gemini variables in
[Choosing the AI Provider](#choosing-the-ai-provider).

## Choosing the AI Provider

Every AI call in the application funnels through a single wrapper,
`utils.call_openai_api()`. That wrapper picks the provider from one environment
variable:

| `USE_GEMINI` | Provider used |
|---|---|
| unset / `false` (default) | OpenAI chat-completions — unchanged behaviour, needs `OPENAI_API_KEY` |
| `true` / `1` / `yes` / `on` | Google Gemini via [`gemini_client.py`](gemini_client.py) — needs Google credentials, **no OpenAI key required** |

Nothing else in the application changes: the same prompts, temperatures, seeds,
JSON contracts and return shapes are used for both providers.

### Option A — Gemini on Vertex AI (service account)

```bash
export USE_GEMINI=true
export GOOGLE_PROJECT_ID=prj-kb-poc1-geminiai-gcp-1010
export GOOGLE_LOCATION=global
export GENAI_MODEL_NAME=gemini-3.1-flash-lite
export GOOGLE_APPLICATION_CREDENTIALS=/path/to/service-account.json

export DATABASE_URL=postgresql://myuser:mypassword@localhost:5432/karmayogi_db

streamlit run app.py --server.port 5000
```

### Option B — Gemini Developer API (API key)

```bash
export USE_GEMINI=true
export GEMINI_API_KEY=...            # used only when GOOGLE_PROJECT_ID is unset
export GENAI_MODEL_NAME=gemini-3.1-flash-lite
```

If you keep these in a `.env` file, load it before starting Streamlit — the app
reads plain environment variables and does not parse `.env` itself:

```bash
set -a && source .env && set +a
streamlit run app.py --server.port 5000
```

### Environment variables

| Variable | Applies to | Default | Purpose |
|---|---|---|---|
| `USE_GEMINI` | both | *(unset → OpenAI)* | `true`/`1`/`yes`/`on` switches the app to Gemini |
| `OPENAI_API_KEY` | OpenAI | — | Required in OpenAI mode only; the client is built lazily, so it is not needed in Gemini mode |
| `GOOGLE_PROJECT_ID` | Gemini | — | Vertex AI project (alias: `GOOGLE_CLOUD_PROJECT`). Set → Vertex AI auth is used |
| `GOOGLE_LOCATION` | Gemini | `global` | Vertex AI location (alias: `GOOGLE_CLOUD_LOCATION`) |
| `GOOGLE_APPLICATION_CREDENTIALS` | Gemini | — | Path to the service-account JSON (read by `google-auth`) |
| `GEMINI_API_KEY` / `GOOGLE_API_KEY` | Gemini | — | Gemini Developer API key; used when `GOOGLE_PROJECT_ID` is unset |
| `GENAI_MODEL_NAME` | Gemini | `gemini-3.1-flash-lite` | Gemini model id |
| `GENAI_MAX_OUTPUT_TOKENS` | Gemini | `8192` | Output-token cap (higher than OpenAI's 2000 because Gemini counts thinking tokens against it) |
| `GENAI_THINKING_LEVEL` | Gemini | `low` | `low`, `high`, or anything else to leave the model default |
| `DATABASE_URL` | both | SQLite fallback | PostgreSQL connection string |
| `SESSION_SECRET` | both | — | Session signing secret |

### Verifying which provider is live

On startup the application logs exactly one provider line — check it first when
debugging a deployment. It never echoes a key or credential contents:

```
[AI provider] Gemini | model=gemini-3.1-flash-lite | auth=Vertex AI (project=my-project, location=global, credentials=set) | max_output_tokens=8192 | thinking_level=low
[AI provider] OpenAI | model=gpt-4o (default) | OPENAI_API_KEY=set
```

A misconfigured Gemini deployment logs `auth=NO CREDENTIALS FOUND` at startup,
and AI calls then fail with `Error calling Gemini API: ...`. Provider errors are
always attributed to the provider that actually served the request — Gemini
traffic is never reported as OpenAI. Note that the internal wrapper function is
still named `call_openai_api()` for backwards compatibility; it is the
provider-agnostic entry point, not an indication that OpenAI was called.

### Deployment (Docker)

The image is built by [`build.sh`](build.sh) / [`Jenkinsfile`](Jenkinsfile) and
needs no change to run on Gemini — the provider is chosen purely by environment
variables at container start.

```bash
# Gemini on Vertex AI: pass the env vars and mount the service-account JSON
docker run -p 5000:5000 \
  -e USE_GEMINI=true \
  -e GOOGLE_PROJECT_ID=prj-kb-poc1-geminiai-gcp-1010 \
  -e GOOGLE_LOCATION=global \
  -e GENAI_MODEL_NAME=gemini-3.1-flash-lite \
  -e GOOGLE_APPLICATION_CREDENTIALS=/secrets/gcp-sa.json \
  -e DATABASE_URL=postgresql://user:pass@host:5432/karmayogi_db \
  -e SESSION_SECRET=... \
  -v /host/path/gcp-sa.json:/secrets/gcp-sa.json:ro \
  <org>/analyser-cbc-service:<tag>
```

Deployment checklist:

1. `GOOGLE_APPLICATION_CREDENTIALS` must point to a path **inside** the
   container; mount the service-account JSON read-only. Never bake it into the
   image or commit it.
2. The service account needs the Vertex AI user role (`roles/aiplatform.user`)
   on the project.
3. On GKE/Cloud Run, omit the credentials variable and file entirely and use
   Workload Identity / the attached service account instead — `google-auth`
   picks those up automatically.
4. `OPENAI_API_KEY` can be left unset in Gemini mode; the OpenAI client is
   never constructed.
5. Confirm the `[AI provider]` startup line in the container logs before
   sign-off.

### Provider differences handled by the wrapper

The two SDKs differ in a few places; `gemini_client.py` normalises all of them
so callers see identical inputs and outputs:

| Concern | OpenAI | Gemini | Handling |
|---|---|---|---|
| JSON mode | `response_format={"type": "json_object"}` | `response_mime_type="application/json"` | Translated per call |
| Output cap | `max_tokens=2000` | `max_output_tokens` (also consumed by thinking tokens) | Defaults to 8192, `GENAI_MAX_OUTPUT_TOKENS` overrides |
| Model id | `gpt-4o` | `gemini-*` | OpenAI-style ids passed by callers are mapped to `GENAI_MODEL_NAME`; explicit Gemini ids pass through |
| Determinism | `seed` | `seed` | Passed through unchanged (temperature/seed values are identical for both) |
| Reasoning/thinking | n/a | thinking tokens can consume the whole budget | `GENAI_THINKING_LEVEL` (default `low`); an empty `MAX_TOKENS` response is retried once with a doubled budget |
| Response text | `choices[0].message.content` | `response.text`, which can be `None` | Candidate parts are walked, thought parts skipped |
| JSON output shape | always a JSON object | may add a markdown fence, a preamble, or a top-level array | Fences/preambles are stripped; a non-object falls back to the existing parse-failure shape |
| Safety blocks | HTTP error | empty response + `block_reason`/`finish_reason` | Turned into the standard `Error calling Gemini API: ...` string |
| Errors | `Error calling OpenAI API: ...` | `Error calling Gemini API: ...` | Both detected by `utils.is_api_error_message()` |

## Repository Layout

```
app.py                    # Main Streamlit app — UI, wizard flows, auth, orchestration
case_generator.py         # Case Study Generator pipeline: prompts, section drafting, continuity, compliance
assessment_criteria.py    # CBC-India AGK Review Rubric definition + scoring logic
utils.py                  # AI provider wrapper (call_openai_api), document extraction, CaseConnect analysis, helpers
gemini_client.py          # Gemini provider (Vertex AI / Developer API); used only when USE_GEMINI is set
db_models.py              # SQLAlchemy models: users, assessment_history, generated_cases
scripts/                  # Utility scripts
docs/DEVELOPER.md         # Full developer & architecture documentation (integration handover)
```

## Documentation

- **[docs/DEVELOPER.md](docs/DEVELOPER.md)** — developer documentation prepared for the AGK Platform integration review: AI engine & model architecture, prompt management, tech stack, async/job handling, infrastructure requirements, security & compliance, dependencies & licensing, and database schema.

## Status

This repository is a working prototype. For AGK Platform integration, the recommended approach is to port the prompt library, wizard flows, and AI pipeline logic into the platform's native stack rather than embedding the Streamlit UI as-is.
