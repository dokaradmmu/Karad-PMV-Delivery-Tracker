"""
GitHub-backed persistence for the RTS (Returned to Sender) exclusion list.

The RTS list is the single authoritative, user-curated record of which
PMV Application Numbers have been confirmed as Returned to Sender. It is
stored as a simple CSV in a GitHub repo so it persists across app restarts
and across whichever machine/session the Streamlit app happens to be
running on (Streamlit Cloud's local filesystem is not reliably persistent
between deploys).

Auth: a GitHub Personal Access Token with "repo" (or fine-grained
"contents: read & write") scope, provided via Streamlit secrets:

    # .streamlit/secrets.toml
    [github]
    token = "ghp_xxx..."
    repo = "dokaradmmu/PMV-Delivery-Tracker"   # owner/repo
    branch = "main"
    file_path = "data/rts_exclusion_list.csv"

If secrets aren't configured yet (e.g. first local test run before the repo
exists), functions degrade to an in-memory / local-file fallback so the app
still works for development - see `_LOCAL_FALLBACK_PATH`.
"""

from __future__ import annotations

import base64
import io
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

APP_NO_COL = "PMV Application Number"
_LOCAL_FALLBACK_PATH = Path("rts_exclusion_list_local.csv")


def _github_config() -> dict | None:
    try:
        cfg = st.secrets["github"]
        return {
            "token": cfg["token"],
            "repo": cfg["repo"],
            "branch": cfg.get("branch", "main"),
            "file_path": cfg.get("file_path", "data/rts_exclusion_list.csv"),
        }
    except Exception:
        return None


def _api_url(cfg: dict) -> str:
    return f"https://api.github.com/repos/{cfg['repo']}/contents/{cfg['file_path']}"


def _headers(cfg: dict) -> dict:
    return {
        "Authorization": f"token {cfg['token']}",
        "Accept": "application/vnd.github+json",
    }


def load_rts_list() -> tuple[pd.DataFrame, str | None]:
    """
    Returns (dataframe with one column 'PMV Application Number', sha).
    `sha` is GitHub's blob sha for the current file version, needed to
    commit an update; it's None when using the local fallback or when the
    file doesn't exist on GitHub yet (first-ever save will create it).
    """
    cfg = _github_config()
    if cfg is None:
        if _LOCAL_FALLBACK_PATH.exists():
            return pd.read_csv(_LOCAL_FALLBACK_PATH, dtype=str), None
        return pd.DataFrame({APP_NO_COL: []}), None

    resp = requests.get(_api_url(cfg), headers=_headers(cfg), params={"ref": cfg["branch"]})
    if resp.status_code == 404:
        return pd.DataFrame({APP_NO_COL: []}), None
    resp.raise_for_status()
    payload = resp.json()
    content = base64.b64decode(payload["content"]).decode("utf-8")
    df = pd.read_csv(io.StringIO(content), dtype=str)
    return df, payload["sha"]


def save_rts_list(app_numbers: set[str], sha: str | None, commit_message: str) -> str | None:
    """
    Overwrite the RTS list on GitHub with the given full set of Application
    Numbers. Returns the new sha (for chaining further updates in the same
    session), or None if using the local fallback.
    """
    df = pd.DataFrame({APP_NO_COL: sorted(app_numbers)})
    csv_bytes = df.to_csv(index=False).encode("utf-8")

    cfg = _github_config()
    if cfg is None:
        df.to_csv(_LOCAL_FALLBACK_PATH, index=False)
        return None

    body = {
        "message": commit_message,
        "content": base64.b64encode(csv_bytes).decode("utf-8"),
        "branch": cfg["branch"],
    }
    if sha:
        body["sha"] = sha

    resp = requests.put(_api_url(cfg), headers=_headers(cfg), json=body)
    resp.raise_for_status()
    return resp.json()["content"]["sha"]
