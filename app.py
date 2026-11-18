"""
Asycuda Tools -- web version (Streamlit).
Run karanna:  streamlit run app.py
"""
import hmac
import io
import os
import time
from datetime import datetime

import streamlit as st

st.set_page_config(page_title="Asycuda Tools", page_icon="🚢", layout="wide")

SECRET_KEYS = ["APP_PASSWORD", "GMAIL_USER", "GMAIL_APP_PASS", "TARGET_EMAIL",
               "BREVO_LOGIN_USER", "BREVO_SMTP_KEY", "SENDER_EMAIL",
               "SUPABASE_URL", "SUPABASE_KEY", "SUPABASE_BUCKET",
               "GITHUB_TOKEN", "GITHUB_BOT_REPO"]


def S(name, default=""):
    try:
        v = st.secrets.get(name, None)
    except Exception:
        v = None
    if v is None:
        v = os.environ.get(name, default)
    return str(v) if v is not None else default


for _k in SECRET_KEYS:                       # core scripts os.environ walin kiyawanawa
    _v = S(_k)
    if _v:
        os.environ[_k] = _v

from core.paths import WORK_DIR                       # noqa: E402
from jobs import JobManager                           # noqa: E402
from runner import Runner, DATA_KEY, EXP_KEY, PDF_PREFIX, MERGED_PREFIX, BACKUP_PREFIX  # noqa: E402
from storage import LocalBackend, Store, SupabaseBackend, dec, enc  # noqa: E402
import viewer                                         # noqa: E402

REQUIRED_SHEETS = ["CUSDEC DETAILS", "CDN DETAILS", "SUMMARY", "b1", "b2", "b3", "b4", "b5", "ex cus", "ex cdn"]
VIEW_DIR = "/tmp/asycuda_view"


# ------------------------------------------------------------------ resources
@st.cache_resource
def get_jobs():
    return JobManager()


@st.cache_resource
def get_store():
    url, key = S("SUPABASE_URL"), S("SUPABASE_KEY")
    if url and key:
        be = SupabaseBackend(url, key, S("SUPABASE_BUCKET", "asycuda"))
        be.ensure_bucket()
        return Store(be, WORK_DIR), False
    return Store(LocalBackend("/tmp/asycuda_storage"), WORK_DIR), True


jobs = get_jobs()
store, LOCAL_ONLY = get_store()
runner = Runner(store, jobs)


# ------------------------------------------------------------------ password gate
def gate():
    pw = S("APP_PASSWORD")
    if not pw:
        st.error("APP_PASSWORD secret eka set karala naha. Streamlit Settings > Secrets eke APP_PASSWORD = \"...\" danna.")
        st.stop()
    if st.session_state.get("auth_ok"):
        return
    st.title("🚢 Asycuda Tools")
    with st.form("login"):
        entered = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            if hmac.compare_digest(entered.encode(), pw.encode()):
                st.session_state["auth_ok"] = True
                st.rerun()
            else:
                st.error("Password waradi.")
    st.stop()


gate()

# ------------------------------------------------------------------ helpers
def need(*names):
    missing = [n for n in names if not S(n)]
    if missing:
        st.error("Secrets walin me ewa naha: " + ", ".join(missing))
        return False
    return True


def mail_env():
    return {"GMAIL_USER": S("GMAIL_USER"), "GMAIL_APP_PASS": S("GMAIL_APP_PASS"),
            "TARGET_EMAIL": S("TARGET_EMAIL", "bathiyapradeep@yahoo.com")}


def brevo_env():
    return {"BREVO_LOGIN_USER": S("BREVO_LOGIN_USER"), "BREVO_SMTP_KEY": S("BREVO_SMTP_KEY"),
            "SENDER_EMAIL": S("SENDER_EMAIL", "export@bastellogistics.com")}


def start(name, fn):
    try:
        jobs.start(name, fn)
    except RuntimeError as e:
        st.error(str(e))
        return
    st.rerun()


def busy():
    return jobs.running()


def fmt_size(n):
    n = float(n or 0)
    for u in ("B", "KB", "MB", "GB"):
        if n < 1024 or u == "GB":
            return f"{n:.0f} {u}" if u == "B" else f"{n:.1f} {u}"
        n /= 1024


# ------------------------------------------------------------------ header + job panel
st.title("🚢 Asycuda Tools")
if LOCAL_ONLY:
    st.warning("⚠ SUPABASE_URL / SUPABASE_KEY set karala naha -> TEST mode (data restart unama ain wenawa). Real use walata Supabase set karanna.")


def job_body():
    j = jobs.current
    if j is None:
        st.caption("Job ekak naha. Yata tabs walin ekak start karanna.")
        return
    icon = {"running": "⏳", "done": "✅", "error": "❌", "cancelled": "⏹"}[j.status]
    took = ((j.finished or datetime.now()) - j.started).seconds
    st.markdown(f"**{icon} {j.name}** — {j.status} ({took}s)")
    if j.status == "running":
        if st.button("⏹ Stop", key="stopbtn"):
            jobs.stop()
    st.code(j.text() or "...", language=None)
    seen = st.session_state.get("seen_status")
    if seen != (id(j), j.status):
        st.session_state["seen_status"] = (id(j), j.status)
        if j.status != "running" and seen is not None:
            st.rerun()          # full rerun -> buttons enable wenawa


with st.container(border=True):
    st.fragment(run_every=2 if busy() else None)(job_body)()

tab_run, tab_chk, tab_exp, tab_data, tab_files, tab_bot = st.tabs(
    ["▶ Run", "🔎 Checks", "✉ Exporters", "📊 Data", "📁 Files", "🤖 Telegram"])

# ------------------------------------------------------------------ RUN
with tab_run:
    st.subheader("Mail → Boat → Merge → Mail")
    c1, c2 = st.columns(2)
    mismatch = c1.selectbox("CDN row count samana nathnam (boat)",
                            ["Skip karanna (safe)", "Ignore karala PDF hadanna"], disabled=busy())
    overwrite = c2.checkbox("PDF kalin thiyenawanam uda liyanna", value=True, disabled=busy())
    on_mismatch = "skip" if mismatch.startswith("Skip") else "go"
    confirm = st.checkbox("Mail eka yawanna ok (mail yawana step walata)", value=False, disabled=busy())

    if st.button("🚀 Run ALL (extract → boat → merge → mail)", type="primary", disabled=busy() or not confirm,
                 use_container_width=True):
        if need("GMAIL_USER", "GMAIL_APP_PASS"):
            start("Run ALL", lambda: runner.run_all(mail_env(), on_mismatch, overwrite))
    if not confirm:
        st.caption("Run ALL / Send mail walata uda 'Mail eka yawanna ok' tick karanna.")

    st.divider()
    b1, b2, b3, b4 = st.columns(4)
    if b1.button("1️⃣ Extract mail", disabled=busy(), use_container_width=True):
        if need("GMAIL_USER", "GMAIL_APP_PASS"):
            start("Extract mail", lambda: runner.extract_mail(mail_env()))
    if b2.button("2️⃣ Boat (PDF)", disabled=busy(), use_container_width=True):
        start("Boat", lambda: runner.boat(on_mismatch, overwrite))
    if b3.button("3️⃣ Merge", disabled=busy(), use_container_width=True):
        start("Merge", runner.merge)
    if b4.button("4️⃣ Send mail", disabled=busy() or not confirm, use_container_width=True):
        if need("GMAIL_USER", "GMAIL_APP_PASS"):
            start("Send mail", lambda: runner.send_mail(mail_env()))

# ------------------------------------------------------------------ CHECKS
with tab_chk:
    st.subheader("Export release check (customs)")
    t1, t2 = st.columns(2)
    test1 = t1.checkbox("TEST mode (palamu rows tikak, Excel wenas karanne naha)", key="t1", disabled=busy())
    rows1 = t2.number_input("Test rows", 1, 50, 3, key="r1", disabled=busy())
    if st.button("▶ Export check", disabled=busy()):
        start("Export check", lambda: runner.export_check(test1, int(rows1)))

    st.divider()
    st.subheader("Trico SN (boat note) check")
    st.caption("Browser eken tricologi site ekata login wela, PHPSESSID cookie eka copy karala daanna (expire unoth aluth ekak).")
    sess = st.text_input("PHPSESSID", type="password", disabled=busy())
    t3, t4 = st.columns(2)
    test2 = t3.checkbox("TEST mode", key="t2", disabled=busy())
    rows2 = t4.number_input("Test rows", 1, 50, 3, key="r2", disabled=busy())
    if st.button("▶ Trico check", disabled=busy()):
        if not sess.strip():
            st.error("PHPSESSID eka danna.")
        else:
            start("Trico check", lambda: runner.trico_check(sess, test2, int(rows2)))

# ------------------------------------------------------------------ EXPORTERS MAIL
with tab_exp:
    st.subheader("Exporters walata mail (Brevo)")
    if not store.has(EXP_KEY):
        st.warning("Exporters Excel eka (CDA_Exporters_Directory.xlsx) upload karala naha. 'Files' tab eken upload karanna.")
    test3 = st.checkbox("TEST mode (Row 2 witharak yawanawa)", value=True, disabled=busy())
    start_row = st.number_input("Start row (TEST mode off nam)", 2, 100000, 2, disabled=busy() or test3)
    ok3 = True if test3 else st.checkbox("Hari, okkoma exporters lata yawanna (Stop button eken nawaththanna puluwan)", disabled=busy())
    if st.button("✉ Send", disabled=busy() or not ok3):
        if need("BREVO_LOGIN_USER", "BREVO_SMTP_KEY"):
            start("Exporters mail", lambda: runner.mail_exporters(brevo_env(), test3, int(start_row)))

# ------------------------------------------------------------------ DATA VIEWER
with tab_data:
    st.subheader("Excel data (laptop / phone)")
    os.makedirs(VIEW_DIR, exist_ok=True)
    view_path = os.path.join(VIEW_DIR, "data.xlsx")
    rc1, rc2 = st.columns([1, 3])
    if rc1.button("🔄 Refresh", disabled=False) or not os.path.exists(view_path):
        try:
            if not store.pull_file(DATA_KEY, view_path):
                st.info("data.xlsx storage eke naha. 'Files' tab eken upload karanna.")
        except Exception as e:
            st.error(f"Pull fail: {e}")
    if os.path.exists(view_path):
        try:
            mtime = os.path.getmtime(view_path)
            if st.session_state.get("wb_mtime") != mtime:
                st.session_state["wb"] = viewer.load(view_path)
                st.session_state["wb_mtime"] = mtime
            wb = st.session_state["wb"]
            allsh = rc2.checkbox("Sheets okkoma pennanna", value=False)
            names = viewer.sheet_names(wb, allsh)
            sheet = st.selectbox("Sheet", names)
            q = st.text_input("Search (row eke ona text ekak)", "")
            ws = wb[sheet]
            rows = viewer.search_rows(ws, q)
            body = rows[1:] if rows and rows[0] == 1 else rows
            per = 100
            pages = max(1, (len(body) + per - 1) // per)
            pg = st.number_input(f"Page (1-{pages})", 1, pages, 1)
            show = [1] + body[(pg - 1) * per: pg * per] if ws.max_row >= 1 else []
            st.caption(f"Rows: {len(body)}  |  Sheet: {sheet}")
            st.markdown(viewer.render(ws, wb, show), unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Excel eka pennanna bari una: {e}")

# ------------------------------------------------------------------ FILES
with tab_files:
    st.subheader("data.xlsx")
    if busy():
        st.info("Job ekak run wenawa -> upload/delete disable (job iwara wenakan inna).")

    up = st.file_uploader("Aluth data.xlsx upload (palaweni eka replace wenawa, backup ekak thiyenawa)", type=["xlsx"],
                          key="up_data", disabled=busy())
    if up is not None and st.button("💾 data.xlsx save karanna", disabled=busy()):
        import openpyxl
        raw = up.getvalue()
        try:
            wbt = openpyxl.load_workbook(io.BytesIO(raw))
            miss = [s for s in REQUIRED_SHEETS if s not in wbt.sheetnames]
        except Exception as e:
            st.error(f"Excel file eka open karanna bari una: {e}")
            miss = None
        if miss:
            st.error("Me sheets naha: " + ", ".join(miss))
        elif miss is not None:
            tmp = os.path.join(WORK_DIR, "_upload.xlsx")
            with open(tmp, "wb") as f:
                f.write(raw)
            runner._backup_remote_data()
            store.push_file(DATA_KEY, tmp)
            os.remove(tmp)
            st.session_state.pop("wb_mtime", None)
            st.success("data.xlsx save una (palaweni eka backups ekata giya).")

    if store.has(DATA_KEY):
        if st.button("⬇ Latest data.xlsx ready karanna"):
            p = os.path.join(WORK_DIR, "_dl.xlsx")
            store.pull_file(DATA_KEY, p)
            st.session_state["dl_data"] = open(p, "rb").read()
        if st.session_state.get("dl_data"):
            st.download_button("⬇ data.xlsx download", st.session_state["dl_data"], file_name="data.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    bk = sorted(store.b.list(BACKUP_PREFIX))
    if bk:
        with st.expander(f"Backups ({len(bk)})"):
            sel = st.selectbox("Backup", bk[::-1])
            if st.button("Me backup eka download ready karanna"):
                p = os.path.join(WORK_DIR, "_bk_dl.xlsx")
                store.b.download(sel, p)
                st.session_state["dl_bk"] = (sel, open(p, "rb").read())
            if st.session_state.get("dl_bk"):
                st.download_button("⬇ " + st.session_state["dl_bk"][0].split("/")[-1], st.session_state["dl_bk"][1],
                                   file_name=st.session_state["dl_bk"][0].split("/")[-1])

    st.divider()
    st.subheader("Exporters Excel (CDA_Exporters_Directory.xlsx)")
    up2 = st.file_uploader("Exporters Excel upload (sheet name: Exporters)", type=["xlsx"], key="up_exp", disabled=busy())
    if up2 is not None and st.button("💾 Exporters Excel save karanna", disabled=busy()):
        import openpyxl
        try:
            wbe = openpyxl.load_workbook(io.BytesIO(up2.getvalue()), read_only=True)
            ok = "Exporters" in wbe.sheetnames
        except Exception as e:
            st.error(f"Open karanna bari una: {e}")
            ok = None
        if ok is False:
            st.error("'Exporters' kiyana sheet eka naha.")
        elif ok:
            tmp = os.path.join(WORK_DIR, "_exp_upload.xlsx")
            with open(tmp, "wb") as f:
                f.write(up2.getvalue())
            store.push_file(EXP_KEY, tmp)
            os.remove(tmp)
            st.success("Save una.")
    st.caption("Exporters file eka: " + ("✅ thiyenawa" if store.has(EXP_KEY) else "❌ naha"))

    st.divider()
    st.subheader("PDF files (storage)")
    for label, prefix in (("pdf/ (extract + boat walin hadapu)", PDF_PREFIX), ("merged/ (mail yawanna kalin)", MERGED_PREFIX)):
        items = store.b.list(prefix)
        with st.expander(f"{label} — {len(items)} files"):
            for key, size in sorted(items.items()):
                name = dec(key.rsplit("/", 1)[1])
                c1, c2 = st.columns([4, 1])
                c1.write(f"{name}  ({fmt_size(size)})")
                if c2.button("⬇", key=f"dlb_{key}"):
                    p = os.path.join(WORK_DIR, "_f.pdf")
                    store.b.download(key, p)
                    st.session_state["dl_pdf"] = (name, open(p, "rb").read())
            if st.session_state.get("dl_pdf"):
                st.download_button("⬇ " + st.session_state["dl_pdf"][0], st.session_state["dl_pdf"][1],
                                   file_name=st.session_state["dl_pdf"][0], mime="application/pdf", key=f"dlp_{prefix}")
            if items:
                ok = st.checkbox(f"{prefix}/ okkoma delete karanna (undo naha)", key=f"cb_{prefix}", disabled=busy())
                if st.button(f"🗑 {prefix}/ clear", key=f"clr_{prefix}", disabled=busy() or not ok):
                    for k in items:
                        store.b.delete(k)
                    st.rerun()

# ------------------------------------------------------------------ TELEGRAM BOT
with tab_bot:
    st.subheader("Telegram bot (on / off)")
    st.caption("Bot eka GitHub Actions (free) eke run wenawa. On karama max ~5h 50m run wela, GitHub walin auto off wenawa.")
    if not (S("GITHUB_TOKEN") and S("GITHUB_BOT_REPO")):
        st.info("GITHUB_TOKEN ekai GITHUB_BOT_REPO (eg: username/telegram-bot) ekai secrets walata danna.")
    else:
        from ghbot import BotControl
        bc = BotControl(S("GITHUB_TOKEN"), S("GITHUB_BOT_REPO"))
        try:
            runs = bc.active_runs()
            st.markdown("Status: " + ("🟢 **ON**" if runs else "⚪ **OFF**"))
        except Exception as e:
            runs = []
            st.error(f"Status ganna bari una: {e}")
        b1, b2 = st.columns(2)
        if b1.button("▶ ON karanna", disabled=bool(runs)):
            try:
                st.success("Start una." if bc.start() == "started" else "Kalinma ON.")
            except Exception as e:
                st.error(str(e))
        if b2.button("⏹ OFF karanna", disabled=not runs):
            try:
                st.success(f"Stop kala ({bc.stop()}).")
            except Exception as e:
                st.error(str(e))
