"""
Keeps passwords and API keys in config.yaml encrypted.

On Windows they are encrypted with your Windows login (DPAPI, the same thing
Chrome and Windows Credential Manager use). The file then holds e.g.
    password: "enc:AQAAANCMnd8BFdERjHoAwE..."
which only your Windows account on this computer can turn back into the
password. Copied to another laptop or user, the values are useless.

Encrypted fields: every `password`, `app_password` and `api_key` in config.yaml.
"""

import base64
import os
import sys

import yaml

PREFIX = "enc:"
SECRET_KEYS = {"password", "app_password", "api_key"}
_ENTROPY = b"JobBot config v1"
FAILED = []   # fields that couldn't be decrypted (config copied from another computer/user)


if sys.platform == "win32":
    import ctypes
    from ctypes import wintypes

    class _Blob(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    _crypt32, _kernel32 = ctypes.windll.crypt32, ctypes.windll.kernel32
    _kernel32.LocalFree.argtypes = [ctypes.c_void_p]

    def _blob(data):
        buf = ctypes.create_string_buffer(data, len(data))
        return _Blob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char))), buf

    def _call(fn, data):
        src, _keep = _blob(data)
        ent, _keep2 = _blob(_ENTROPY)
        out = _Blob()
        if not fn(ctypes.byref(src), None, ctypes.byref(ent), None, None, 0x1, ctypes.byref(out)):
            raise OSError("Windows couldn't encrypt/decrypt this value")
        try:
            return ctypes.string_at(out.pbData, out.cbData)
        finally:
            _kernel32.LocalFree(out.pbData)

    def _encrypt(data):
        return _call(lambda s, d, e, r, p, f, o: _crypt32.CryptProtectData(s, d, e, r, p, f, o), data)

    def _decrypt(data):
        return _call(lambda s, d, e, r, p, f, o: _crypt32.CryptUnprotectData(s, None, e, r, p, f, o), data)

    AVAILABLE = True
else:   # other systems: stored as before (plain text)
    AVAILABLE = False


def encrypt(value):
    value = "" if value is None else str(value)
    if not AVAILABLE or not value or value.startswith(PREFIX):
        return value
    return PREFIX + base64.b64encode(_encrypt(value.encode("utf-8"))).decode("ascii")


def decrypt(value, where=""):
    value = "" if value is None else str(value)
    if not value.startswith(PREFIX):
        return value
    try:
        return _decrypt(base64.b64decode(value[len(PREFIX):])).decode("utf-8")
    except Exception:
        FAILED.append(where or "?")
        return ""


def _walk(cfg, fn, path=""):
    changed = False
    for k, v in list((cfg or {}).items()):
        here = f"{path}{k}"
        if isinstance(v, dict):
            changed |= _walk(v, fn, here + ".")
        elif str(k) in SECRET_KEYS and v not in (None, ""):
            new = fn(v, here)
            if new != v:
                cfg[k], changed = new, True
    return changed


def protect(cfg):
    """Encrypt every plain-text secret in cfg (in place). Returns True if anything changed."""
    return _walk(cfg, lambda v, _: encrypt(v))


def reveal(cfg):
    """Decrypt every secret in cfg (in place) for the bot to use. Returns cfg."""
    _walk(cfg, decrypt)
    return cfg


def secret_paths(cfg):
    out = []
    _walk(cfg, lambda v, p: out.append(p) or v)
    return out


def protect_file(path, header=None):
    """Encrypt the secrets inside a YAML file on disk. Returns True if it was rewritten."""
    if not (AVAILABLE and os.path.isfile(path)):
        return False
    with open(path, encoding="utf-8") as f:
        text = f.read()
    cfg = yaml.safe_load(text) or {}
    if not isinstance(cfg, dict) or not protect(cfg):
        return False
    if header is None:
        header = "".join(l + "\n" for l in text.splitlines() if l.startswith("#")) + "\n"
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(header)
        yaml.safe_dump(cfg, f, allow_unicode=True, sort_keys=False, width=120)
    os.replace(tmp, path)
    return True
