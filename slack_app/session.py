"""In-memory session store keyed by the parent message's ts (== thread_ts
for replies in that thread). Good enough for a single always-running
process; a real database would only be needed for multi-instance/horizontal
scaling, which a small internal tool doesn't need."""

import threading


class Session:
    def __init__(self, channel, params, kind="sample_buy"):
        self.channel = channel
        self.params = params
        self.kind = kind  # "sample_buy" (rate cards -> workbook) or "strata_order" (workbook -> .sbx)
        self.file_paths = []
        self.file_names = []

    def add_file(self, path, name):
        self.file_paths.append(path)
        self.file_names.append(name)


class SessionStore:
    def __init__(self):
        self._sessions = {}
        self._lock = threading.Lock()

    def create(self, thread_ts, channel, params, kind="sample_buy"):
        with self._lock:
            self._sessions[thread_ts] = Session(channel, params, kind=kind)
        return self._sessions[thread_ts]

    def get(self, thread_ts):
        with self._lock:
            return self._sessions.get(thread_ts)

    def discard(self, thread_ts):
        with self._lock:
            self._sessions.pop(thread_ts, None)


sessions = SessionStore()
