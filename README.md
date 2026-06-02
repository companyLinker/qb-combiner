# QB Combiner — Streamlit Web App

Internal tool to consolidate per-entity QuickBooks P&L + Balance Sheet exports into a SUMIFS-linked combination workbook.

**Self-hosted.** Single password. No data leaves your server.

## Quick start (local)

```bash
# 1. Set your password
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# edit secrets.toml and set APP_PASSWORD

# 2. Run with Docker
docker compose up -d --build

# 3. Open
open http://localhost:8501
```

For VM hosting (cloud), see **[DEPLOYMENT.md](./DEPLOYMENT.md)**.

## What it does

A four-step wizard:

1. **📂 Upload Files** — drag in your QB exports (.xlsx, multiple) + optional target combination template
2. **📊 Variants & Analysis** — see consolidated data, browse the chart-of-accounts variants, download the master + digest workbooks
3. **🔗 Review Mapping** — auto-mapping handles ~99% of P&L and ~93% of BS automatically; review the rest in an editable table
4. **💾 Generate Linked Workbook** — produces the final SUMIFS-linked combination workbook you can download

## Folder layout

```
qb_combiner_app/
├── app.py                    # Entry point (landing page + auth)
├── pages/
│   ├── 1_Upload_Files.py
│   ├── 2_Variants_and_Analysis.py
│   ├── 3_Review_Mapping.py
│   └── 4_Generate_Linked_Workbook.py
├── lib/
│   ├── parser.py             # QB Excel parser
│   ├── master_builder.py     # Master + digest builders
│   ├── linked_builder.py     # SUMIFS workbook builder
│   └── mapping_rules.py      # Auto-mapping rule engine
├── .streamlit/
│   ├── config.toml           # Streamlit server config
│   └── secrets.toml.example  # Copy → secrets.toml, set password
├── Dockerfile
├── docker-compose.yml
├── requirements.txt
└── DEPLOYMENT.md
```

## Run without Docker (Python dev)

```bash
python -m venv .venv
source .venv/bin/activate    # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
streamlit run app.py
```

## Privacy & security notes

- **No persistent storage.** Files live in browser session memory only. Closing the tab = data gone.
- **Single-password gate.** Read from `.streamlit/secrets.toml`. Not committed to git.
- **No telemetry.** Streamlit's usage stats are disabled via `config.toml`.
- For multi-user authentication (named accounts, OAuth, etc.), wrap behind an auth proxy like Authelia or Cloudflare Access.

## See also

- `../skill_quickbooks_combiner/` — the underlying Python skill (CLI scripts + mapping rules)
- `../skill_quickbooks_combiner/SKILL.md` — the design doc / Claude skill manifest
