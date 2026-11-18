"""
Hama step ekakatama: storage eken pull -> original script (core/*) run -> storage ekata push.
Scripts wala logic wenas karala naha (boat witharai LibreOffice walata wenas kale).
"""
import importlib
import os
import shutil
import time
from datetime import datetime

from core import paths
from core.paths import (DATA_XLSX, EXPORTERS_XLSX, LOGS_DIR, MERGED_DIR, PDF_DIR, RECALC_DIR, WORK_DIR)

DATA_KEY = "data.xlsx"
EXP_KEY = "exporters.xlsx"
PDF_PREFIX = "pdf"
MERGED_PREFIX = "merged"
BACKUP_PREFIX = "backups"
KEEP_BACKUPS = 10


class Runner:
    def __init__(self, store, jobs):
        self.store = store
        self.jobs = jobs

    # ------------------------------------------------------------ sync helpers
    def _pull_data(self):
        paths.ensure_dirs()
        if not self.store.pull_file(DATA_KEY, DATA_XLSX):
            raise RuntimeError("Storage eke data.xlsx naha. 'Files' tab eken data.xlsx upload karanna.")

    def _backup_remote_data(self):
        """data.xlsx wenas karanna kalin backup ekak (palaweni 10 ma witharai)."""
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        tmp = os.path.join(WORK_DIR, "_bk.xlsx")
        if self.store.pull_file(DATA_KEY, tmp):
            self.store.push_file(f"{BACKUP_PREFIX}/data_{stamp}.xlsx", tmp)
            os.remove(tmp)
            keys = sorted(self.store.b.list(BACKUP_PREFIX))
            for k in keys[:-KEEP_BACKUPS]:
                self.store.b.delete(k)

    def _push_data_if_changed(self, before_md5):
        from storage import file_md5
        if os.path.exists(DATA_XLSX) and file_md5(DATA_XLSX) != before_md5:
            self._backup_remote_data()
            self.store.push_file(DATA_KEY, DATA_XLSX)
            print("data.xlsx storage ekata save kala.")

    def _mod(self, name):
        return importlib.import_module(f"core.{name}")

    def _with_data(self, fn, dirs=()):
        """data.xlsx (+ pdf/merged folders) pull -> fn() -> (finally) push."""
        from storage import file_md5
        self._pull_data()
        before = file_md5(DATA_XLSX)
        snaps = [(prefix, local, self.store.pull_dir(prefix, local)) for prefix, local in dirs]
        try:
            fn()
        finally:
            self._push_data_if_changed(before)
            for prefix, local, snap in snaps:
                self.store.push_dir(prefix, local, snap)

    # ------------------------------------------------------------ steps
    def extract_mail(self, env):
        os.environ.update(env)
        m = self._mod("main1")
        m.GMAIL_USER, m.GMAIL_APP_PASS = env["GMAIL_USER"], env["GMAIL_APP_PASS"]
        self._with_data(m.main_process, dirs=[(PDF_PREFIX, PDF_DIR)])

    def boat(self, on_mismatch="skip", overwrite=True):
        b = self._mod("boat")
        self._with_data(lambda: b.process_and_pdf(on_mismatch=on_mismatch, overwrite=overwrite, log=self.jobs.log),
                        dirs=[(PDF_PREFIX, PDF_DIR)])

    def merge(self):
        m = self._mod("merge")
        shutil.rmtree(PDF_DIR, ignore_errors=True)
        shutil.rmtree(MERGED_DIR, ignore_errors=True)
        os.makedirs(PDF_DIR); os.makedirs(MERGED_DIR)
        s1 = self.store.pull_dir(PDF_PREFIX, PDF_DIR)
        s2 = self.store.pull_dir(MERGED_PREFIX, MERGED_DIR)
        try:
            m.merge_pdfs()
        finally:
            self.store.push_dir(PDF_PREFIX, PDF_DIR, s1)
            self.store.push_dir(MERGED_PREFIX, MERGED_DIR, s2)

    def send_mail(self, env):
        from core import office
        os.environ.update(env)
        m = self._mod("mail")
        m.GMAIL_USER, m.GMAIL_APP_PASS = env["GMAIL_USER"], env["GMAIL_APP_PASS"]
        m.TARGET_EMAIL = env.get("TARGET_EMAIL", m.TARGET_EMAIL)
        self._pull_data()
        shutil.rmtree(MERGED_DIR, ignore_errors=True)
        os.makedirs(MERGED_DIR)
        snap = self.store.pull_dir(MERGED_PREFIX, MERGED_DIR)
        # mail.py: SUMMARY!AA3/AA4 (formula) -> LibreOffice walin recalculate karapu copy eken kiyawanawa
        shutil.rmtree(RECALC_DIR, ignore_errors=True)
        recalced = office.recalculated_copy(DATA_XLSX, RECALC_DIR)
        m.EXCEL_FILE = recalced
        try:
            m.send_email_with_attachments()
        finally:
            # mail.py yawapu pasu merged PDF delete karanawa -> storage eka ekatama samana karanawa
            self.store.push_dir(MERGED_PREFIX, MERGED_DIR, snap)

    def run_all(self, env, on_mismatch="skip", overwrite=True):
        """extract -> boat -> merge -> mail   (run_all.bat ekatama samana; ekak fail unoth nawaththanawa)"""
        print("=== STEP 1/4 : Extract mail ===")
        self.extract_mail(env)
        self.jobs.check_cancel()
        print("\n=== STEP 2/4 : Boat ===")
        self.boat(on_mismatch, overwrite)
        self.jobs.check_cancel()
        print("\n=== STEP 3/4 : Merge ===")
        self.merge()
        self.jobs.check_cancel()
        print("\n=== STEP 4/4 : Send mail ===")
        self.send_mail(env)
        print("\n=== ALL STEPS COMPLETED ===")

    def export_check(self, test_mode=False, test_rows=3, delay=0.3):
        m = self._mod("export")
        m.time = self.jobs.time_proxy
        m.TEST_MODE, m.TEST_ROWS, m.DELAY = test_mode, test_rows, delay
        self._with_data(m.run_automation)

    def trico_check(self, phpsessid, test_mode=False, test_rows=3, delay=0.3):
        m = self._mod("trico")
        m.time = self.jobs.time_proxy
        m.PHPSESSID = phpsessid.strip()
        m.TEST_MODE, m.TEST_ROWS, m.DELAY = test_mode, test_rows, delay
        self._with_data(m.run_automation)

    def mail_exporters(self, env, test_mode=True, start_row=2):
        os.environ.update(env)
        m = self._mod("mailexporters")
        m.time = self.jobs.time_proxy
        m.BREVO_LOGIN_USER = env["BREVO_LOGIN_USER"]
        m.SENDER_PASSWORD = env["BREVO_SMTP_KEY"]
        m.SENDER_EMAIL = env.get("SENDER_EMAIL", m.SENDER_EMAIL)
        m.TEST_MODE, m.START_ROW = test_mode, int(start_row)
        if not self.store.pull_file(EXP_KEY, EXPORTERS_XLSX):
            raise RuntimeError("Exporters Excel eka (CDA_Exporters_Directory.xlsx) upload karala naha. 'Files' tab eken upload karanna.")
        m.send_emails()
