# SpinCycle

> An offline-capable data cleaning and QC pipeline for analytical chemistry instrument exports (Agilent Chemstation, MassHunter).

[![CI](https://github.com/stefanhaugen/spincycle_sanitized/actions/workflows/ci.yml/badge.svg)](https://github.com/stefanhaugen/spincycle_sanitized/actions/workflows/ci.yml)
[![Python](https://img.shields.io/badge/python-3.11-blue.svg)](https://www.python.org/downloads/release/python-3119/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

---

## The problem

Analytical chemistry labs export raw HPLC/MS results from instrument software (Agilent ChemStation, MassHunter) as Excel files that are inconsistently formatted, mix metadata with data, and require hours of manual cleanup before QC review including replicate averaging, calibration curve validation, LLOQ/ULOQ flagging, and CVS (continuing verification standard) checks.

Instrument PCs are typically locked-down (no admin rights, no internet) and run a single Windows OS image for years. Installing Python or third-party tools is not an option.

**SpinCycle solves this by shipping as a self-contained, fully offline folder.** A one-time setup on any internet-connected machine bundles an embedded Python 3.11 distribution + all dependencies into the folder. After setup, the folder can be copied via USB or network share to any Windows machine  (including standalone instrument PCs)  and the Streamlit pipeline launches in the user's default browser with a double-click.

## Features

- **Two parse modes**: ChemStation (Data + Labels sheets) and MassHunter (Sheet1 with `*Results` headers), auto-detected per file.
- **Standards workup** with configurable LLOQ/ULOQ bounds, automatic linearity flagging, and recovery validation.
- **CVS (Continuing Verification Standard) QC** with per-analyte tolerance assignment, bulk-action editing, and pass/fail heatmaps.
- **Batch processing** — drag in multiple files, get one consolidated Excel + per-file sheets.
- **Export formats**: multi-sheet `.xlsx` (workup, summary, tidy, raw) and SQLite databases for downstream analysis.
- **Optional Dropbox push** to a configurable group folder, with automatic filename collision handling.
- **Fully offline-capable** — the entire dependency graph (`streamlit`, `pandas`, `numpy`, `openpyxl`, `xlsxwriter`, `xlrd`, `dropbox`) is bundled as local wheels.

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                     User's Windows PC                           │
│                                                                 │
│  launch_140SpinCycle.bat                                        │
│         │                                                       │
│         └──> python/python.exe  (embedded 3.11.9, in-folder)    │
│                    │                                            │
│                    └──> streamlit run app.py                    │
│                              │                                  │
│                              ├─ Parsers (Chemstation/MassHunter)│
│                              ├─ QC engine (Standards + CVS)     │
│                              ├─ Excel/SQLite exporters          │
│                              └─ Dropbox client (optional)       │
│                                                                 │
│  Browser opens → http://localhost:8501                          │
└─────────────────────────────────────────────────────────────────┘
```

The app is a single Streamlit process that runs locally. There is no server component, no database (SQLite outputs are user-downloadable, not persistent), and no network calls except optional outbound Dropbox uploads.

## Quick start

### For lab users (Windows, offline-target)

1. Copy the entire `spincycle_sanitized` folder onto an internet-connected machine.
2. Double-click `SETUP_once.bat`. This downloads Python 3.11.9 and installs all dependencies into the folder.
3. Copy the now-self-contained folder to the target machine (USB or network share).
4. Double-click `launch_140SpinCycle.bat`. The app opens in your default browser.

See [docs for lab users](README_bestplacetostart.txt) for the original, plain-language guide.

### For developers (any OS)

```bash
git clone https://github.com/stefanhaugen/spincycle_sanitized.git
cd spincycle_sanitized
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
pip install -r requirements-dev.txt  # tests + linters
streamlit run app.py
```

App is available at <http://localhost:8501>.

## Configuration

SpinCycle reads configuration from Streamlit's `secrets.toml` (preferred) or session state (transient).

Create `.streamlit/secrets.toml`:

```toml
# Optional — enables Dropbox push UI
dropbox_token = "sl.YOUR_TOKEN_HERE"
dropbox_dest  = "/Analytical"
```

| Setting | Default | Purpose |
|---|---|---|
| `dropbox_token` | unset | Dropbox API token; without it, the push UI is hidden. |
| `dropbox_dest`  | `/Analytical` | Target folder for uploaded reports. URLs and unprefixed paths are normalized. |

**Never commit `secrets.toml` to the repo** — it is excluded via `.gitignore`.

## Development workflow

```bash
# Run linters + formatters (also enforced by pre-commit)
ruff check .
black --check .

# Run tests
pytest -v

# Run tests with coverage
pytest --cov=. --cov-report=term-missing
```

Pre-commit hooks run `black`, `ruff`, and `detect-secrets` automatically on every commit:

```bash
pre-commit install
```

See [CONTRIBUTING.md](CONTRIBUTING.md) for the full PR workflow.

## Deployment

### Offline-bundled (current production model)

`SETUP_once.bat` produces a self-contained folder; copy to target. No further configuration needed.

### Docker (developer / cloud)

```bash
docker build -t spincycle:latest .
docker run --rm -p 8501:8501 spincycle:latest
```

See [Dockerfile](Dockerfile) for image details.

## Troubleshooting

For end-user errors (Windows offline deployment), see the [RUNBOOK](RUNBOOK.md).

Common developer issues:

- **`ImportError: pyarrow` on pandas 3.0+** — `app.py` sets `pd.set_option("future.infer_string", False)` at startup; if this fails, downgrade pandas to <3.0 in `requirements.txt`.
- **Streamlit cache not invalidating after edits** — bump `_CACHE_VER` in `app.py`.
- **Dropbox token errors** — verify token scope includes `files.content.write` and the token has not expired.

## Project status

Currently in active development. The core ChemStation/MassHunter pipeline is production-stable. Planned work:

- Multi-file consolidated reporting (in progress)
- Statistical analyses page (placeholder exists)
- HPLC main-page dashboard (placeholder exists)

## License

MIT — see [LICENSE](LICENSE).

## Author

**Stefan Haugen** — [github.com/stefanhaugen](https://github.com/stefanhaugen)
