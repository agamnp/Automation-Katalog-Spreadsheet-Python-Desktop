# globals.py
_stop_requested = False

def set_stop_requested(value: bool):
    global _stop_requested
    _stop_requested = value

def get_stop_requested() -> bool:
    return _stop_requested