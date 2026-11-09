"""Per-person scan activity and cooperative cancellation, local to each thread."""
from contextlib import contextmanager
from contextvars import ContextVar


class ScanCancelled(Exception):
    pass


_context = ContextVar('face_scan_context', default=(None, None))


@contextmanager
def scan_context(activity=None, cancel_check=None):
    token = _context.set((activity, cancel_check))
    try:
        check_scan_cancelled()
        yield
    finally:
        _context.reset(token)


def check_scan_cancelled():
    _, cancel_check = _context.get()
    if cancel_check and cancel_check():
        raise ScanCancelled('Đã dừng quét theo yêu cầu')


def report_scan_activity(stage, **details):
    check_scan_cancelled()
    callback, _ = _context.get()
    if callback:
        callback({'stage': stage, **details})
