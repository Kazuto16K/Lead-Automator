"""Keep Local_Lead_Tracker.xlsx in a (separate, private) GitHub repo so it survives Streamlit Cloud restarts.

- On app start: download the latest file from GitHub (it is the source of truth).
- After every save or delete: upload the file back (one commit per change = full history).
Configure with Streamlit secrets; if they are missing the app simply works locally with no sync.

    GITHUB_TOKEN  = "github_pat_..."      # fine-grained token: Contents read+write on the data repo ONLY
    GITHUB_REPO   = "your-user/lead-data"  # must NOT be the repo Streamlit deploys from (a push there redeploys the app)
    GITHUB_BRANCH = "main"                 # optional
    GITHUB_PATH   = "Local_Lead_Tracker.xlsx"  # optional
"""
import base64
from pathlib import Path

import requests

API = "https://api.github.com"
TIMEOUT = 20


def config(secrets):
    """Settings dict from a secrets mapping, or None when sync is not configured."""
    try:
        token, repo = secrets.get("GITHUB_TOKEN"), secrets.get("GITHUB_REPO")
        if not token or not repo:
            return None
        return {"token": token, "repo": repo, "branch": secrets.get("GITHUB_BRANCH") or "main",
                "path": secrets.get("GITHUB_PATH") or "Local_Lead_Tracker.xlsx"}
    except Exception:  # no secrets file at all (local run)
        return None


def _headers(cfg, accept="application/vnd.github+json"):
    return {"Authorization": f"Bearer {cfg['token']}", "Accept": accept, "X-GitHub-Api-Version": "2022-11-28"}


def _url(cfg):
    return f"{API}/repos/{cfg['repo']}/contents/{cfg['path']}"


def pull(cfg, local_path):
    """Download the file from GitHub over local_path. Returns 'pulled', 'missing' (not in repo yet) or 'error: ...'."""
    try:
        r = requests.get(_url(cfg), params={"ref": cfg["branch"]},
                         headers=_headers(cfg, "application/vnd.github.raw+json"), timeout=TIMEOUT)
        if r.status_code in (404, 409):  # not there yet (404), or the repo has no commits at all (409)
            return "missing"
        r.raise_for_status()
        Path(local_path).write_bytes(r.content)
        return "pulled"
    except Exception as e:  # noqa: BLE001 - never crash the app over a sync problem
        return f"error: {e}"


def push(cfg, local_path, message):
    """Upload local_path to GitHub as a new commit. Returns (ok, text)."""
    try:
        r = requests.get(_url(cfg), params={"ref": cfg["branch"]}, headers=_headers(cfg), timeout=TIMEOUT)
        sha = r.json().get("sha") if r.status_code == 200 else None
        if r.status_code not in (200, 404, 409):
            r.raise_for_status()
        body = {"message": message, "branch": cfg["branch"],
                "content": base64.b64encode(Path(local_path).read_bytes()).decode()}
        if sha:
            body["sha"] = sha
        r = requests.put(_url(cfg), json=body, headers=_headers(cfg), timeout=TIMEOUT)
        r.raise_for_status()
        return True, "backed up to GitHub"
    except Exception as e:  # noqa: BLE001
        return False, f"GitHub backup failed: {e}"
