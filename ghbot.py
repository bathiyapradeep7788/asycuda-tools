"""Telegram bot eka GitHub Actions walin on/off karanna (workflow_dispatch + cancel)."""
import requests


class BotControl:
    def __init__(self, token, repo, workflow="bot.yml", ref="main", api="https://api.github.com"):
        self.repo, self.workflow, self.ref, self.api = repo, workflow, ref, api.rstrip("/")
        self.h = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json",
                  "X-GitHub-Api-Version": "2022-11-28"}

    def _req(self, method, path, **kw):
        r = requests.request(method, f"{self.api}{path}", headers=self.h, timeout=30, **kw)
        return r

    def active_runs(self):
        out = []
        for st in ("in_progress", "queued"):
            r = self._req("GET", f"/repos/{self.repo}/actions/workflows/{self.workflow}/runs", params={"status": st, "per_page": 10})
            if r.status_code != 200:
                raise RuntimeError(f"GitHub {r.status_code}: {r.text[:200]}")
            out += r.json().get("workflow_runs", [])
        return out

    def start(self):
        if self.active_runs():
            return "already"
        r = self._req("POST", f"/repos/{self.repo}/actions/workflows/{self.workflow}/dispatches", json={"ref": self.ref})
        if r.status_code != 204:
            raise RuntimeError(f"Start fail: {r.status_code} {r.text[:200]}")
        return "started"

    def stop(self):
        runs = self.active_runs()
        for run in runs:
            r = self._req("POST", f"/repos/{self.repo}/actions/runs/{run['id']}/cancel")
            if r.status_code not in (202, 409):
                raise RuntimeError(f"Stop fail: {r.status_code} {r.text[:200]}")
        return len(runs)
