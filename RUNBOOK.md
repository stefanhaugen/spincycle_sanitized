# SpinCycle Operations Runbook

> Diagnostic and recovery procedures for common SpinCycle failures. Audience: lab admins and operators running the offline-bundled deployment, plus developers debugging issues.

This runbook complements the user-facing troubleshooting section in the [README](README.md) — that one covers normal user issues, this one covers the deeper failures and their root causes.

## Severity definitions

| Sev | Definition | Example | Response time |
|---|---|---|---|
| 1 | App won't start, no users can run pipeline | `python.exe` missing, all packages corrupted | Immediate |
| 2 | App starts but core pipeline is broken | ChemStation parser raises on every file | Same day |
| 3 | Partial degradation — some features broken | Dropbox push fails, rest works | This week |
| 4 | Cosmetic or annoyance | UI label wrong, slow but functional | Backlog |

---

## Quick reference table

| Symptom | Likely cause | Section |
|---|---|---|
| "Embedded Python not found" on launch | `python/` folder missing | [§ Embedded Python missing](#embedded-python-missing) |
| Import errors in browser tab on launch | Stale or corrupted packages | [§ Package corruption](#package-corruption) |
| Streamlit launches but shows blank page | Cache invalidation needed | [§ Streamlit cache stuck](#streamlit-cache-stuck) |
| "Could not find pip wheel" during SETUP | Tampered or deleted `packages/` folder | [§ Missing wheel cache](#missing-wheel-cache) |
| ChemStation/MassHunter parse fails | Unexpected sheet format from instrument | [§ Parser failure](#parser-failure) |
| Dropbox push silently fails | Expired token or wrong scope | [§ Dropbox auth failures](#dropbox-auth-failures) |
| "Maximum recursion depth" or memory error | Pathological input file (huge batch) | [§ Resource exhaustion](#resource-exhaustion) |
| CI failing on push | Lint/format/test regression | [§ CI failures](#ci-failures) |

---

## Embedded Python missing

**Severity:** 1
**Symptom:** Double-clicking `launch_140SpinCycle.bat` shows `[ERROR] Embedded Python not found!` and exits.

**Diagnosis**

The `python/` folder inside the spincycle folder is missing or incomplete. This happens when:

- The folder was copied to a target machine without first running `SETUP_once.bat` on an internet-connected machine.
- Someone manually deleted `python/` (it can look like clutter).
- An antivirus quarantined files in `python/`.

**Verify**

Run from the spincycle folder:
```cmd
dir python\python.exe
```

If "File Not Found", the embedded interpreter is missing.

**Recovery**

1. On an **internet-connected** machine, in the spincycle folder, double-click `SETUP_once.bat`.
2. Wait for "SETUP COMPLETE!" message (~2 minutes).
3. Copy the now-rebuilt folder (containing populated `python/`) back to the target machine.

The setup script is idempotent — safe to re-run.

---

## Package corruption

**Severity:** 1
**Symptom:** App launches but the browser shows an `ImportError` or `ModuleNotFoundError`.

**Diagnosis**

The `python/Lib/site-packages` directory has corrupted or partial installs. Common triggers: power loss during SETUP, full disk during install, antivirus removed files mid-install.

**Verify**

```cmd
python\python.exe -c "import streamlit, pandas, numpy, openpyxl, xlsxwriter, xlrd, dropbox; print('all imports OK')"
```

If any module fails to import, packages are corrupted.

**Recovery**

1. Delete the `python/` folder entirely.
2. Re-run `SETUP_once.bat` on an internet-connected machine.
3. Copy back to target if needed.

Do not try to `pip install` selectively — the bundled wheel cache in `packages/` is what SETUP uses for offline installs. Rebuilding the whole `python/` folder is the supported recovery path.

---

## Streamlit cache stuck

**Severity:** 3
**Symptom:** App launches but shows stale data, blank widgets, or repeats a previous computation that shouldn't apply.

**Diagnosis**

`@st.cache_data` decorators in `app.py` are still serving cached results from a previous session or code version. The `_CACHE_VER = "v12"` constant in `app.py` is the version key — incrementing it invalidates all caches.

**Recovery (user-facing)**

In the running app: hold **Shift** and click the browser refresh button (forces a hard refresh). If that doesn't work, close the browser tab and re-launch.

**Recovery (developer-facing)**

After code changes that affect cached function behavior, bump `_CACHE_VER` in `app.py`:

```python
_CACHE_VER = "v13"   # was v12, bumped after parser change
```

Commit alongside the actual change. This forces every cache to invalidate on next run.

---

## Missing wheel cache

**Severity:** 1
**Symptom:** During `SETUP_once.bat`: `[ERROR] Could not find pip wheel in packages folder.` or `[ERROR] Package installation failed.`

**Diagnosis**

The `packages/` folder is missing required `.whl` files. SpinCycle ships pre-downloaded wheels for fully-offline installation. If someone deleted `packages/` or copied an incomplete folder, SETUP can't proceed.

**Verify**

```cmd
dir packages\*.whl
```

You should see ~10–15 `.whl` files (pip, setuptools, streamlit, pandas, numpy, openpyxl, xlsxwriter, xlrd, dropbox, and their transitive deps).

**Recovery**

You need a complete spincycle folder from a known-good source. Options:

1. Re-clone from GitHub: `git clone https://github.com/stefanhaugen/spincycle_sanitized.git` — but note the public sanitized repo may not include the `packages/` folder; check internal artifacts.
2. Copy `packages/` from another working SpinCycle installation.
3. Rebuild from `requirements.txt`: on an internet-connected machine, run `python -m pip download -r requirements.txt -d packages/` to repopulate.

---

## Parser failure

**Severity:** 2
**Symptom:** Uploading a file returns "Could not parse submission" or a Python traceback in the UI.

**Diagnosis**

The instrument produced an Excel file with an unexpected sheet structure. Common causes:

- ChemStation export missing the "Labels" or "Data" sheet.
- MassHunter export missing the `*Results` column suffixes.
- A user manually edited the export file and broke the layout.
- A new instrument firmware version changed the export schema.

**Verify**

Open the failing Excel file in Excel itself. Check that:

- ChemStation: there are sheets named exactly "Data" and "Labels".
- MassHunter: there is a sheet whose row 1 contains columns ending in `Results` (e.g. `Compound 1 Results`).

**Recovery**

If a single file is malformed → re-export from the instrument software. If the parser is broken across all files of a given type, the export schema may have changed: open an issue with one sample file attached and the SpinCycle developer will update the parser.

---

## Dropbox auth failures

**Severity:** 3
**Symptom:** "Dropbox upload failed: ..." in the UI, or the Dropbox push button does nothing.

**Diagnosis**

The Dropbox API token configured in `.streamlit/secrets.toml` (or session state) is invalid. Possible causes:

- Token expired (Dropbox tokens have explicit expiration when created via OAuth flow).
- Token revoked by user/admin in Dropbox account settings.
- Token has wrong scope — needs `files.content.write` minimum.
- Network policy on instrument PC blocking `api.dropbox.com`.

**Verify**

```bash
curl -X POST https://api.dropboxapi.com/2/users/get_current_account \
  -H "Authorization: Bearer YOUR_TOKEN_HERE"
```

A valid token returns JSON with account info. An invalid token returns `401 Unauthorized`.

**Recovery**

1. Visit `https://www.dropbox.com/developers/apps` and either generate a new token on your existing app or create a new app.
2. Ensure the app has the `files.content.write` permission.
3. Update `.streamlit/secrets.toml`:
   ```toml
   dropbox_token = "sl.NEW_TOKEN_HERE"
   ```
4. Restart the Streamlit server.

Tokens are sensitive — never commit `secrets.toml` to the repo (it's in `.gitignore`).

---

## Resource exhaustion

**Severity:** 2
**Symptom:** App hangs, browser shows "loading..." indefinitely, or Python process crashes with `MemoryError`.

**Diagnosis**

Batch mode processing too many files at once, or single file with extreme row count (>500K rows). Instrument PCs typically have 4–8 GB RAM; pandas can exceed that on large batches.

**Verify**

Open Task Manager → Performance tab → watch memory consumption while uploading. If Python process exceeds 80% of RAM, it's resource-bound.

**Recovery**

Short term: process files in smaller batches (5–10 at a time instead of 50+). Long term: move to a more powerful machine, or implement chunked reading in the pipeline.

---

## CI failures

**Severity:** 4 (CI failures don't affect production users, only future commits)
**Symptom:** Red X on GitHub Actions runs, badge in README shows failing.

**Diagnosis**

1. Click the failed run, find the first red step.
2. Common causes: black/ruff formatting drift, new test failure, secret detected in a committed file.

**Recovery (formatting drift)**

```bash
pip install -r requirements-dev.txt
pre-commit run --all-files     # auto-fixes formatting
git add -A && git commit -m "Fix formatting" --no-verify
git push
```

**Recovery (test failure)**

Run locally first: `pytest -v`. The failure message points to the offending test. Either fix the code or update the test if requirements legitimately changed.

**Recovery (detected secret)**

If the secret is a false positive (e.g., a placeholder in an example):
```bash
detect-secrets scan --update .secrets.baseline
git add .secrets.baseline
git commit -m "Update secrets baseline"
```

If the secret is real, **rotate it immediately**, then remove from history (this is more involved — search "git remove secret from history" for the BFG Repo-Cleaner workflow).

---

## Escalation

For issues not covered here, contact: **stefanhaugen@gmail.com** (or open a GitHub issue with the "operational" label).

For Sev 1 outages: capture the failing terminal output and the contents of `python/` folder listing (`dir python /s > listing.txt`) before any recovery attempts — that gives the dev team forensic data to root-cause the issue.

---

## Runbook conventions

This runbook follows three principles:

1. **One section per failure mode** — each section is self-contained and can be linked directly.
2. **Verify before recover** — diagnostic command first, recovery second. Don't blindly rebuild.
3. **Preserve forensic data** — when something breaks at Sev 1, capture state before fixing if possible.

When adding a new section, copy the template from another section: symptom, severity, diagnosis, verify, recovery.
