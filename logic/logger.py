# -*- coding: utf-8 -*-
"""
Application-wide error logger with deduplication.

Identical errors (same type + message + traceback) are recorded only once
with a repeat counter.  Logs live in data/error.log so they survive
alongside the database.
"""
import sys
import os
import traceback
import hashlib
import json
import datetime
import threading
import logging as _logging

_APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG_DIR = os.path.join(_APP_DIR, "data")
LOG_FILE = os.path.join(LOG_DIR, "error.log")
SEEN_FILE = os.path.join(LOG_DIR, "_error_seen.json")
MAX_LOG_SIZE = 512 * 1024  # 512 KB before rotation

_lock = threading.Lock()
_seen = {}
_loaded = False


def _ensure_dir():
    if not os.path.isdir(LOG_DIR):
        try:
            os.makedirs(LOG_DIR)
        except Exception:
            pass


def _load_seen():
    global _seen, _loaded
    if _loaded:
        return
    _loaded = True
    _ensure_dir()
    if os.path.isfile(SEEN_FILE):
        try:
            with open(SEEN_FILE, "r") as f:
                _seen = json.load(f)
        except Exception:
            _seen = {}


def _save_seen():
    _ensure_dir()
    try:
        with open(SEEN_FILE, "w") as f:
            json.dump(_seen, f, ensure_ascii=False)
    except Exception:
        pass


def _rotate_if_needed():
    if os.path.isfile(LOG_FILE) and os.path.getsize(LOG_FILE) > MAX_LOG_SIZE:
        bak = LOG_FILE + ".old"
        try:
            os.replace(LOG_FILE, bak)
        except Exception:
            pass


def _hash_error(etype, value, tb_lines):
    """Produce a short hash for deduplication."""
    raw = "\n".join([etype, str(value)] + tb_lines[-8:])  # last 8 frames
    return hashlib.md5(raw.encode("utf-8", errors="replace")).hexdigest()[:12]


def _write(msg):
    _ensure_dir()
    _rotate_if_needed()
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass


def log_error(etype, value, tb, extra_info=""):
    """Log an error.  Deduplicates by type+message+traceback hash."""
    _load_seen()
    tb_lines = traceback.format_exception(etype, value, tb)

    # Special handling: if it's a chained exception, include __context__ too
    if value and getattr(value, "__context__", None) and value.__context__ is not value:
        tb_lines.append("\n-- caused by --\n")
        tb_lines.extend(traceback.format_exception(
            type(value.__context__), value.__context__,
            value.__context__.__traceback__))

    h = _hash_error(str(etype), str(value), tb_lines)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with _lock:
        if h in _seen:
            _seen[h]["count"] += 1
            _seen[h]["last"] = now
            _save_seen()
            return  # suppress duplicate

        _seen[h] = {"count": 1, "first": now, "last": now}
        _save_seen()

        lines = []
        lines.append("=" * 72)
        lines.append("FIRST:  {}".format(now))
        lines.append("HASH:   {}".format(h))
        if extra_info:
            lines.append("INFO:   {}".format(extra_info))
        lines.append("-" * 72)
        for line in tb_lines:
            lines.append(line.rstrip("\n"))
        lines.append("")

        _write("\n".join(lines))


def log_warning(msg):
    """Write a plain warning/note (not deduplicated)."""
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _write("[WARN  {}] {}".format(now, msg))


def log_info(msg):
    """Write a plain informational message."""
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _write("[INFO  {}] {}".format(now, msg))


def get_log_path():
    return LOG_FILE


def get_recent_errors(n=20):
    """Return the last N lines of the error log."""
    if not os.path.isfile(LOG_FILE):
        return ""
    try:
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            lines = f.readlines()
        return "".join(lines[-n:])
    except Exception:
        return ""


def install_hook():
    """Replace sys.excepthook so all uncaught exceptions get logged."""
    _orig = sys.excepthook

    def _hook(etype, value, tb):
        try:
            log_error(etype, value, tb, extra_info="uncaught exception")
        except Exception:
            pass
        if _orig:
            _orig(etype, value, tb)
        else:
            sys.__excepthook__(etype, value, tb)

    sys.excepthook = _hook


def install_thread_hook():
    """Patch threading.Thread.run so thread crashes are also captured."""
    _orig_run = threading.Thread.run

    def _wrapped_run(self, *a, **k):
        try:
            _orig_run(self, *a, **k)
        except Exception:
            etype, value, tb = sys.exc_info()
            try:
                log_error(etype, value, tb,
                          extra_info="thread: {}".format(self.name))
            except Exception:
                pass
            raise

    threading.Thread.run = _wrapped_run
