# CBC-India AGK Case Study Suite - Replit Documentation

## Overview

The CBC-India AGK Case Study Suite is a Streamlit-based web application providing three AI-powered tools for case study management and development:

1.  **Case Study Analyser**: Evaluates case studies against the CBC-India AGK Case Study Review Rubric, providing detailed scoring, recommendations, and assessment history. It processes uploaded documents (PDF/DOCX) using the configured AI provider's API (OpenAI or Gemini).
2.  **CaseConnect**: An AI-enabled case discovery tool that recommends AGK repository case studies to faculty based on course outlines and specific criteria (learners, objectives, competencies, duration, sector). It references a CSV database of 62 cases and provides discussion points, key themes, and direct links to the iGOT platform.
3.  **Case Study Generator**: Drafts new case studies from raw source material (transcripts, reports, URLs, notes). It supports Lesson-Drawing, Decision-Forcing, and Caselet case types through an 8-step wizard, including AI source processing, section drafting, compliance review, and optional teaching note generation. Drafts are auto-persisted and exportable in DOCX/PDF/TXT formats.

The application aims to enhance case study quality, facilitate case discovery for educators, and streamline the case study creation process within the CBC-India AGK framework.

## User Preferences

Preferred communication style: Simple, everyday language.

## System Architecture

### Frontend Architecture

The application uses the Streamlit framework with a wide layout. UI components include custom CSS for styling (Poppins font, blue palette #1E3A8A), interactive data visualizations via Plotly, and PDF/DOCX upload capabilities. It follows a single-page application design with session state management for user authentication and data persistence.

### Backend Architecture

The application's core logic is orchestrated by `app.py`. Utility functions for document processing and AI provider integration are in `utils.py` (with the Gemini provider in `gemini_client.py`), while `assessment_criteria.py` defines the rubric and scoring logic. Database operations are handled by `db_models.py`.

The **Assessment Engine** employs a rubric-based evaluation system with four weighted areas, each containing criteria with specific scoring ranges. AI-powered text analysis uses structured prompts sent to the configured AI provider (OpenAI or Gemini) for evaluation, feedback, and score aggregation. The **Document Processing Pipeline** involves extracting text from uploaded PDFs/DOCX files, AI evaluation, score calculation, and PDF report generation.

### Data Storage Solutions

A PostgreSQL database (with SQLite fallback for testing) is used, managed by SQLAlchemy with a declarative base pattern. The schema includes `users` for authentication and `assessment_history` and `generated_cases` for storing assessment results and draft case studies, respectively. Connection management includes pooling, recycling, and timeout handling.

### Authentication and Authorization

User authentication is username/password-based, secured with bcrypt_sha256 hashing. Session state manages user logins, and user registration includes unique username/email constraints.

### AI Integration

The system integrates with one of two interchangeable providers for all AI-powered text analysis: the OpenAI API (GPT models) by default, or Google Gemini (Vertex AI / Gemini Developer API) when the `USE_GEMINI` environment variable is enabled. Every call routes through `utils.call_openai_api()`, which normalises provider differences so the prompts and response shapes are identical either way. Structured prompt engineering is used for each assessment criterion, with JSON response parsing for scores and feedback.

### Reporting and Visualization

Reports are generated using the FPDF library, embedding visualizations from Plotly (interactive charts) and Matplotlib (static charts). A tiered grading system (Excellent, Good, Satisfactory, Needs Improvement) categorizes performance.

## External Dependencies

### Third-Party Services

*   **OpenAI API** (default): Used for AI-powered text analysis, assessment, and content generation. Configured via the `OPENAI_API_KEY` environment variable.
*   **Google Gemini** (optional alternative): Enabled with `USE_GEMINI=true`. Uses Vertex AI when `GOOGLE_PROJECT_ID`, `GOOGLE_LOCATION` and `GOOGLE_APPLICATION_CREDENTIALS` are set, or the Gemini Developer API with `GEMINI_API_KEY`. Model selected via `GENAI_MODEL_NAME` (default `gemini-3.1-flash-lite`). When enabled, no OpenAI key is required.
*   **PostgreSQL Database**: The primary production database, connected via the `DATABASE_URL` environment variable. SQLite is used as a local development fallback.

### Key Python Libraries

*   **Web Framework**: `streamlit`
*   **Database & ORM**: `sqlalchemy`, `psycopg2`
*   **Document Processing**: `PyPDF2`, `python-docx`
*   **Data Analysis & Visualization**: `pandas`, `plotly`, `matplotlib`, `numpy`
*   **AI Integration**: `openai`, `google-genai`
*   **Security**: `passlib`
*   **Report Generation**: `fpdf`
*   **Utilities**: `base64`, `json`, `datetime`

### Environment Configuration

*   **Required**: `DATABASE_URL`, plus provider credentials — `OPENAI_API_KEY` (OpenAI mode) **or** `USE_GEMINI=true` with the Google credentials above (Gemini mode)
*   **Optional (Gemini tuning)**: `GENAI_MODEL_NAME`, `GENAI_MAX_OUTPUT_TOKENS`, `GENAI_THINKING_LEVEL`
*   **Optional (Azure Deployment)**: `AZURE_INFERENCE_SDK_ENDPOINT`, `DEPLOYMENT_NAME`