"""
Persistent storage (data.xlsx, pdf/, merged/) -- Supabase Storage (free) ekak, nathnam local folder ekak (test walata).
Streamlit cloud eke disk eka temporary nisa hama job ekakatama kalin pull, passe push karanawa.
"""
import hashlib
import os
import re
import shutil
import time
from urllib.parse import quote

import requests

SAFE = re.compile(r"[A-Za-z0-9._-]")


def enc(name):
    """Local file name -> storage key (reversible, Supabase key rules walata ok)."""
    out = []
    for ch in name:
        out.append(ch if SAFE.fullmatch(ch) else "".join(f"~{b:02X}" for b in ch.encode("utf-8")))
    return "".join(out)


def dec(key):
    parts, i, raw = [], 0, bytearray()
    while i < len(key):
        if key[i] == "~" and re.fullmatch(r"[0-9A-F]{2}", key[i + 1:i + 3]):
            raw.append(int(key[i + 1:i + 3], 16))
            i += 3
        else:
            if raw:
                parts.append(raw.decode("utf-8", "replace")); raw = bytearray()
            parts.append(key[i]); i += 1
    if raw:
        parts.append(raw.decode("utf-8", "replace"))
    return "".join(parts)


def _join(prefix, name):
    return f"{prefix}/{name}" if prefix else name


def file_md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


class LocalBackend:
    """Test / fallback: folder ekak storage ekak widiyata."""

    def __init__(self, root):
        self.root = root
        os.makedirs(root, exist_ok=True)

    def _p(self, key):
        return os.path.join(self.root, *key.split("/"))

    def list(self, prefix):
        prefix = prefix.rstrip("/")
        d = self._p(prefix) if prefix else self.root
        if not os.path.isdir(d):
            return {}
        return {_join(prefix, n): os.path.getsize(os.path.join(d, n)) for n in os.listdir(d)
                if os.path.isfile(os.path.join(d, n))}

    def upload(self, key, path):
        dst = self._p(key)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(path, dst)

    def download(self, key, path):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        shutil.copy2(self._p(key), path)

    def delete(self, key):
        try:
            os.remove(self._p(key))
        except FileNotFoundError:
            pass

    def exists(self, key):
        return os.path.isfile(self._p(key))

    def size(self, key):
        return os.path.getsize(self._p(key))


class SupabaseBackend:
    def __init__(self, url, key, bucket="asycuda", timeout=60):
        self.base = url.rstrip("/") + "/storage/v1"
        self.bucket = bucket
        self.timeout = timeout
        self.h = {"apikey": key}
        if key.startswith("eyJ"):          # legacy JWT key
            self.h["Authorization"] = f"Bearer {key}"
        self.s = requests.Session()

    def _req(self, method, url, headers=None, **kw):
        hdr = {**self.h, **(headers or {})}
        last = None
        for attempt in range(3):
            try:
                r = self.s.request(method, url, headers=hdr, timeout=self.timeout, **kw)
                if r.status_code >= 500:
                    raise RuntimeError(f"Supabase {r.status_code}: {r.text[:200]}")
                return r
            except (requests.RequestException, RuntimeError) as e:
                last = e
                time.sleep(1.5 * (attempt + 1))
        raise RuntimeError(f"Supabase request fail: {last}")

    def ensure_bucket(self):
        r = self._req("GET", f"{self.base}/bucket/{self.bucket}")
        if r.status_code == 200:
            return
        r = self._req("POST", f"{self.base}/bucket", json={"id": self.bucket, "name": self.bucket, "public": False})
        if r.status_code not in (200, 201) and "already exists" not in r.text.lower():
            raise RuntimeError(f"Bucket hadanna bari una: {r.status_code} {r.text[:200]}")

    def list(self, prefix):
        prefix = prefix.rstrip("/")
        out, offset = {}, 0
        while True:
            r = self._req("POST", f"{self.base}/object/list/{self.bucket}",
                          json={"prefix": prefix, "limit": 100, "offset": offset,
                                "sortBy": {"column": "name", "order": "asc"}})
            if r.status_code != 200:
                raise RuntimeError(f"List fail: {r.status_code} {r.text[:200]}")
            items = r.json()
            for it in items:
                if it.get("id") is None:         # folder
                    continue
                size = (it.get("metadata") or {}).get("size", 0)
                out[_join(prefix, it["name"])] = size
            if len(items) < 100:
                return out
            offset += 100

    def upload(self, key, path):
        with open(path, "rb") as f:
            data = f.read()
        r = self._req("POST", f"{self.base}/object/{self.bucket}/{quote(key)}", data=data,
                      headers={"x-upsert": "true", "Content-Type": "application/octet-stream"})
        if r.status_code not in (200, 201):
            raise RuntimeError(f"Upload fail ({key}): {r.status_code} {r.text[:200]}")

    def download(self, key, path):
        r = self._req("GET", f"{self.base}/object/{self.bucket}/{quote(key)}")
        if r.status_code != 200:
            raise RuntimeError(f"Download fail ({key}): {r.status_code} {r.text[:200]}")
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as f:
            f.write(r.content)

    def delete(self, key):
        r = self._req("DELETE", f"{self.base}/object/{self.bucket}", json={"prefixes": [key]})
        if r.status_code not in (200, 204):
            raise RuntimeError(f"Delete fail ({key}): {r.status_code} {r.text[:200]}")

    def exists(self, key):
        folder, _, _ = key.rpartition("/")
        return key in self.list(folder)

    def size(self, key):
        folder, _, _ = key.rpartition("/")
        return self.list(folder).get(key, 0)


class Store:
    """data.xlsx + pdf/ + merged/ + exporters.xlsx"""

    def __init__(self, backend, work_dir):
        self.b = backend
        self.work = work_dir

    # ---- single files
    def pull_file(self, key, local):
        if not self.b.exists(key):
            if os.path.exists(local):
                os.remove(local)
            return False
        self.b.download(key, local)
        return True

    def push_file(self, key, local):
        self.b.upload(key, local)

    def has(self, key):
        return self.b.exists(key)

    # ---- folders (local folder eka storage prefix ekata samana wenawa)
    @staticmethod
    def snapshot(local_dir):
        snap = {}
        if os.path.isdir(local_dir):
            for n in os.listdir(local_dir):
                p = os.path.join(local_dir, n)
                if os.path.isfile(p):
                    st = os.stat(p)
                    snap[n] = (st.st_size, st.st_mtime_ns)
        return snap

    def pull_dir(self, prefix, local_dir):
        shutil.rmtree(local_dir, ignore_errors=True)
        os.makedirs(local_dir, exist_ok=True)
        for key in self.b.list(prefix):
            self.b.download(key, os.path.join(local_dir, dec(key.rsplit("/", 1)[1])))
        return self.snapshot(local_dir)

    def push_dir(self, prefix, local_dir, snap=None):
        """Storage eka local folder ekata samana karanawa:
        new / wenas unu (size ho mtime) files upload, local nathi ewa storage eken delete."""
        snap = snap or {}
        remote = self.b.list(prefix)
        now = self.snapshot(local_dir)
        keys = {f"{prefix}/{enc(n)}": n for n in now}
        for key, n in keys.items():
            if key not in remote or remote[key] != now[n][0] or snap.get(n) != now[n]:
                self.b.upload(key, os.path.join(local_dir, n))
        for key in remote:
            if key not in keys:
                self.b.delete(key)
