# MCP Office MCP Server

MCP Office is a Model Context Protocol (MCP) server designed to give AI assistants “hands-on” control over Microsoft Office documents and email drafts. Using a single server you can:

- convert Markdown to Word/Excel/PPTX/EML drafts
- edit, reformat, and analyze Word documents with advanced, human-like tools
- store generated files either locally or in your preferred cloud storage backend
- run everything locally or containerize it for easy deployment

---
## Key Features
### Word (DOCX)
Uses advanced tooling for Word, providing dozens of operations:
- Create documents from Markdown templates
- Insert paragraphs, headings, tables, pictures, lists, page breaks, TOCs
- Apply complex formatting (styles, table formatting, cell padding, etc.)
- Manage protection (passwords, restricted editing, digital signatures)
- Fully control footnotes/endnotes with robust validation and repair
- Retrieve or edit comments, convert to PDF, run find-and-replace, etc.

### Excel (XLSX)
- Convert Markdown tables and formulas into real spreadsheets
- Support for relative references (B[0], T1.B[0], etc.)
- Basic styling (bold, italic) and formula propagation

### PowerPoint (PPTX)
- Build presentations from structured slide descriptions
- Choose between 4:3 and 16:9 templates (customizable)
- Automatically remove placeholder slides and populate titles, bullets, authors

### Emails (EML)
- Generate HTML emails with Mustache templates (default or custom)
- Dynamic email templates via `config/email_templates.yaml`
- Enforce strict HTML and formatting guidelines to keep emails consistent

### Storage Backends
All generated files are uploaded through a unified layer (`app/storage`):
- Local filesystem (default)
- AWS S3
- MinIO (S3-compatible)
- Google Cloud Storage
- Azure Blob Storage

The storage layer handles naming, uploads, and optional signed download URLs.

---
## Repository Layout
```
app/
├── config.py            # Centralized configuration + logging
├── main.py              # MCP server bootstrap (FastMCP) and tool registration
├── storage/             # Upload backends (local/S3/MinIO/GCS/Azure)
├── templates/           # Default and custom Office/email templates
├── tools/
│   ├── email/           # Base + dynamic email template tools
│   ├── excel/           # Markdown → XLSX conversion
│   ├── pptx/            # Presentation creation helpers
│   └── word/
│       ├── creation/    # Markdown → Word conversion
│       ├── manipulation/# Advanced Word editing toolbox
│       ├── core/, utils/# Shared primitives and utilities
└── utils/
    └── template_utils.py# Template discovery helpers
```
---
## Requirements
- Python 3.12+
- (Optional) Docker / Docker Compose
- Access to relevant cloud credentials if using non-local storage

---
## Environment Variables
1. Copy `.env.example` to `.env` (the example file enumerates every supported setting and includes inline comments).
2. Pick your desired `UPLOAD_STRATEGY` (`LOCAL`, `S3`, `MINIO`, `GCS`, `AZURE`).
3. Fill in only the variables that apply to that strategy and restart the MCP server.

**Local strategy** works without extra variables—artifacts are written under `output/` (or `/app/output` in Docker). Cloud strategies require the credentials listed below.

| Variable(s) | Description |
| --- | --- |
| `DEBUG` | `true/false` to enable verbose logging |
| `UPLOAD_STRATEGY` | Selects the upload backend |
| `SIGNED_URL_EXPIRES_IN` | Expiration (seconds) for presigned links |
| `AWS_ACCESS_KEY`, `AWS_SECRET_ACCESS_KEY`, `AWS_REGION`, `S3_BUCKET` | AWS S3 credentials and destination bucket |
| `MINIO_ENDPOINT`, `MINIO_ACCESS_KEY`, `MINIO_SECRET_KEY`, `MINIO_BUCKET`, `MINIO_REGION`, `MINIO_VERIFY_SSL`, `MINIO_PATH_STYLE` | MinIO endpoint plus auth and connection toggles |
| `GCS_BUCKET`, `GCS_CREDENTIALS_PATH` | GCS bucket and path to the service-account JSON file (inside the container use `/app/config/...`) |
| `AZURE_STORAGE_ACCOUNT_NAME`, `AZURE_STORAGE_ACCOUNT_KEY`, `AZURE_CONTAINER`, `AZURE_BLOB_ENDPOINT` | Azure Blob credentials; `AZURE_BLOB_ENDPOINT` is optional when using the default cloud hostname |

> For precise validation logic see `app/config.py`.

---
## Local Setup
```bash
git clone https://github.com/Kalagaar/mcp-office.git
cd mcp-office
python -m venv .venv
source .venv/bin/activate          # On Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env               # adjust values to your environment
mkdir -p output custom_templates config
```

Run the server locally:
```bash
python -m app.main
# MCP endpoint exposed at http://localhost:8900/mcp (streamable-http transport)
```

---
## Docker Usage
### Build & Run
```bash
docker build -t mcp-office .
docker run --rm -it \
  -p 8900:8900 \
  --env-file .env \
  -v ${PWD}/output:/app/output \
  -v ${PWD}/custom_templates:/app/custom_templates \
  -v ${PWD}/config:/app/config \
  mcp-office
```

### docker-compose
```bash
docker compose up --build
```
`docker-compose.yml` automatically mounts the key folders and reads `.env`.

---
## Templates
- **Default templates** live under `app/templates/default`.
- **Custom templates** should be placed in `custom_templates/` with well-known names:
  - Word: `custom_docx_template.docx`
  - PowerPoint: `custom_pptx_template_4_3.pptx`, `custom_pptx_template_16_9.pptx`
  - Email: `custom_email_template.html`
- **Dynamic email templates**: place definitions in `config/email_templates.yaml` and refer to HTML files by filename (no path). Each template becomes its own MCP tool.

---
## Available Word Tools (Examples)
Registered via `register_word_manipulation_tools()` in `app/main.py`.
Some highlights:

| Tool | Description |
| --- | --- |
| `word_create_document` | Create DOCX with optional metadata |
| `word_add_paragraph`, `word_add_heading`, `word_add_table`, `word_add_picture` | Insert structured content near text/indices |
| `word_search_replace`, `word_replace_block_between_manual_anchors` | Advanced text manipulation |
| `word_format_text`, `word_set_table_cell_shading`, `word_set_table_column_width` | Fine-grained formatting |
| `word_protect_document`, `word_add_digital_signature`, `word_unprotect_document` | Document protection & signing |
| `word_add_footnote_*`, `word_validate_document_footnotes` | Robust footnote CRUD & validation |
| `word_get_all_comments`, `word_get_comments_by_author` | Comment extraction |
| `word_find_text`, `word_convert_to_pdf` | Search and format conversions |

Refer to `app/main.py` for the full list and parameter descriptions.

---
## Storage Behavior
All tools use `app.storage.upload_file(file, suffix)`, which:
1. Generates a unique object name and writes to the configured backend.
2. For `LOCAL`, saves under `output/` and returns the filesystem path.
3. For cloud targets, returns a presigned URL with expiration `SIGNED_URL_EXPIRES_IN`.

---
## MCP Runtime
- Transport: `streamable-http`
- Default address: `http://0.0.0.0:8900/mcp`
- Configure your MCP-compatible client/assistant with this endpoint and protocol.
