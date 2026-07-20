# PMV Toolkit Delivery Tracker — Karad Division (v1)

Tracks PM Vishwakarma (PMV) toolkit delivery status across two successive
report exports from the PMV website, restricted to Karad Division
jurisdiction, with a persistent Returned-to-Sender (RTS) exclusion list.

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI — uploads, dashboard, RTS list manager, download |
| `pmv_core.py` | Core data-processing pipeline (framework-agnostic, unit-testable) |
| `pmv_excel_writer.py` | Builds the output `.xlsx` workbook from pipeline results |
| `rts_github_store.py` | Reads/writes the persistent RTS list to a GitHub repo |
| `requirements.txt` | Python dependencies |

## What it does

1. You upload the **Office Master File** (Karad's Division/Sub Division/
   Taluka/PIN structure), an **older** PMV report CSV, and a **newer** PMV
   report CSV.
2. It filters both reports down to Karad-jurisdiction PIN codes (everything
   else goes to an "Other Divisions" sheet).
3. It tags each Karad row with Division / Sub Division / Taluka.
4. It diffs old vs new (matched on `PMV Application Number`) to find
   toolkits that moved from Pending to Delivered between the two dates.
5. It applies your persistent RTS exclusion list to split "genuinely stuck"
   Pending items from ones already confirmed as Returned to Sender.
6. Output: an Excel workbook with **Delivered**, **Pending**,
   **Returned to Sender**, **Delivered [old date]–[new date]**, and
   **Other Divisions** sheets. No frozen panes, Arial font throughout.

## Known limitation (by design, v1)

The Office Master File's PIN → Office Name mapping is not 1:1 (multiple
Branch Post Offices commonly share a PIN code), so the output tags
Division / Sub Division / Taluka / PIN but does **not** attempt to guess a
specific Office/BPO name from PIN alone — that would be unreliable.

## RTS exclusion list persistence (GitHub)

The RTS list is the single authoritative record of confirmed
Returned-to-Sender toolkits. It's stored as a CSV in a GitHub repo so it
survives app restarts/redeploys, and is applied automatically on every run.

### One-time setup

1. Create a **private** GitHub repo (e.g. `dokaradmmu/PMV-Delivery-Tracker`).
2. Generate a GitHub Personal Access Token with `repo` scope (classic) or
   `contents: read & write` (fine-grained), scoped to that repo.
3. In Streamlit Cloud, go to your app's **Settings → Secrets** and add:

   ```toml
   [github]
   token = "ghp_xxxxxxxxxxxxxxxxxxxx"
   repo = "dokaradmmu/PMV-Delivery-Tracker"
   branch = "main"
   file_path = "data/rts_exclusion_list.csv"
   ```

4. That's it — the app creates `data/rts_exclusion_list.csv` in the repo
   automatically the first time you save an RTS entry.

### Local development (no GitHub secrets configured yet)

If `st.secrets["github"]` isn't set (e.g. running locally before you've set
up the repo/token), the app automatically falls back to a local file
`rts_exclusion_list_local.csv` in the working directory, so you can still
test the full workflow end-to-end.

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploying

Push this folder to the GitHub repo, then deploy via
[Streamlit Community Cloud](https://streamlit.io/cloud) pointing at
`app.py`, same pattern as your existing COD Digital Payment app. Add the
`[github]` secrets block above under the app's Settings → Secrets.
