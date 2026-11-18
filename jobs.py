"""
Background job runner.
 * Job ekak browser eka close unath (phone sleep unath) wada karagena yanawa (thread ekak).
 * Ekata kalin ekak witharai run wenne (double-click walin mail 2 parak yanna epa).
 * Print karana okkoma log ekata yanawa. 'Stop' button eken cooperative cancel (sleep/log points walin).
"""
import io
import sys
import threading
import time as _time
import traceback
from datetime import datetime


class JobCancelled(BaseException):
    """BaseException -> scripts wala `except Exception` walin gilla gannne naha."""


class _TimeProxy:
    """module.time wenuwata danna: sleep eka cancel check karanawa."""

    def __init__(self, mgr):
        self._mgr = mgr

    def __getattr__(self, name):
        return getattr(_time, name)

    def sleep(self, seconds):
        end = _time.monotonic() + max(0.0, seconds)
        while True:
            self._mgr.check_cancel()
            rem = end - _time.monotonic()
            if rem <= 0:
                return
            _time.sleep(min(0.5, rem))


class _Router(io.TextIOBase):
    """sys.stdout: job thread eke print okkoma job eke log ekata."""

    def __init__(self, mgr, original):
        self._mgr, self._orig = mgr, original

    def write(self, s):
        job = self._mgr._job_of_thread()
        if job is None:
            return self._orig.write(s)
        job._feed(s)
        return len(s)

    def flush(self):
        try:
            self._orig.flush()
        except Exception:
            pass

    def isatty(self):
        return False


class Job:
    def __init__(self, name):
        self.name = name
        self.status = "running"          # running | done | error | cancelled
        self.started = datetime.now()
        self.finished = None
        self.lines = []
        self._partial = ""
        self.cancel = threading.Event()
        self.lock = threading.Lock()

    def _feed(self, s):
        with self.lock:
            self._partial += s
            while "\n" in self._partial:
                line, self._partial = self._partial.split("\n", 1)
                self.lines.append(line)
            if len(self.lines) > 5000:
                self.lines = self.lines[-4000:]

    def text(self, tail=400):
        with self.lock:
            out = self.lines[-tail:] + ([self._partial] if self._partial else [])
        return "\n".join(out)


class JobManager:
    def __init__(self):
        self.current = None
        self.history = []
        self._lock = threading.Lock()
        self._threads = {}
        self.time_proxy = _TimeProxy(self)
        if not isinstance(sys.stdout, _Router):
            sys.stdout = _Router(self, sys.stdout)

    # ---- helpers used inside job threads
    def _job_of_thread(self):
        return self._threads.get(threading.get_ident())

    def check_cancel(self):
        job = self._job_of_thread()
        if job is not None and job.cancel.is_set():
            raise JobCancelled()

    def log(self, *args):
        print(*args)
        self.check_cancel()

    # ---- public
    def running(self):
        j = self.current
        return j is not None and j.status == "running"

    def start(self, name, fn):
        with self._lock:
            if self.running():
                raise RuntimeError(f"'{self.current.name}' dan run wenawa. Eka iwara wenakan inna (nathnam Stop karanna).")
            job = Job(name)
            self.current = job

        def runner():
            self._threads[threading.get_ident()] = job
            try:
                fn()
                final = "done"
            except JobCancelled:
                job._feed("\n*** Stop karanawa (user cancel kala) ***\n")
                final = "cancelled"
            except BaseException as e:           # noqa
                job._feed(f"\n!!! ERROR: {type(e).__name__}: {e}\n{traceback.format_exc()}\n")
                final = "error"
            finally:
                job.finished = datetime.now()
                job._feed("\n")
                job.status = locals().get("final", "error")      # log ekata passe status (ui race nathi wenna)
                self._threads.pop(threading.get_ident(), None)
                self.history.append(job)
                self.history = self.history[-10:]

        t = threading.Thread(target=runner, daemon=True, name=f"job-{name}")
        t.start()
        return job

    def stop(self):
        if self.running():
            self.current.cancel.set()
