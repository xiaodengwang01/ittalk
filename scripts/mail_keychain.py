"""Local macOS prompts and Keychain storage; secrets never enter argv or logs."""
from __future__ import annotations

import ctypes as C
import getpass
import subprocess
import sys


class CredentialError(RuntimeError):
    pass


SERVICE = b"Tailki Alibaba Mail drafts"


def prompt(label, *, secret=False, default="", gui=False):
    if gui:
        hidden = " with hidden answer" if secret else ""
        script = (
            "on run argv\nactivate\n"
            "set answer to display dialog (item 1 of argv) default answer (item 2 of argv)"
            + hidden + ' buttons {"取消", "确定"} default button "确定" cancel button "取消"'
            + ' with title "Tailki · 邮件草稿配置"\nreturn text returned of answer\nend run'
        )
        result = subprocess.run(
            ["/usr/bin/osascript", "-e", script, label, default],
            capture_output=True, text=True,
        )
        if result.returncode:
            raise CredentialError("已取消本机邮箱配置；本地草稿保留。")
        return result.stdout.rstrip("\r\n")
    if not sys.stdin.isatty():
        raise CredentialError("请使用 setup --gui，在本机窗口输入邮箱和专用密码。")
    return getpass.getpass(label + "：") if secret else input(label + "：").strip()


class Keychain:
    def __init__(self):
        if sys.platform != "darwin":
            raise CredentialError("此凭据入口需要 macOS 钥匙串。")
        self.sec = C.CDLL("/System/Library/Frameworks/Security.framework/Security")
        self.cf = C.CDLL("/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation")
        self.cf.CFRelease.argtypes = [C.c_void_p]
        self.sec.SecKeychainFindGenericPassword.argtypes = [
            C.c_void_p, C.c_uint32, C.c_char_p, C.c_uint32, C.c_char_p,
            C.POINTER(C.c_uint32), C.POINTER(C.c_void_p), C.POINTER(C.c_void_p),
        ]
        self.sec.SecKeychainFindGenericPassword.restype = C.c_int32
        self.sec.SecKeychainAddGenericPassword.argtypes = [
            C.c_void_p, C.c_uint32, C.c_char_p, C.c_uint32, C.c_char_p,
            C.c_uint32, C.c_char_p, C.POINTER(C.c_void_p),
        ]
        self.sec.SecKeychainAddGenericPassword.restype = C.c_int32
        self.sec.SecKeychainItemModifyAttributesAndData.argtypes = [
            C.c_void_p, C.c_void_p, C.c_uint32, C.c_char_p,
        ]
        self.sec.SecKeychainItemModifyAttributesAndData.restype = C.c_int32
        self.sec.SecKeychainItemFreeContent.argtypes = [C.c_void_p, C.c_void_p]

    def _find(self, account):
        name = account.encode("utf-8")
        length, data, item = C.c_uint32(), C.c_void_p(), C.c_void_p()
        status = self.sec.SecKeychainFindGenericPassword(
            None, len(SERVICE), SERVICE, len(name), name,
            C.byref(length), C.byref(data), C.byref(item),
        )
        if status == -25300:
            return None, None
        if status:
            raise CredentialError("无法读取本机钥匙串，请检查钥匙串是否已解锁。")
        try:
            secret = C.string_at(data, length.value).decode("utf-8")
        finally:
            self.sec.SecKeychainItemFreeContent(None, data)
        return secret, item

    def get(self, account):
        secret, item = self._find(account)
        if item:
            self.cf.CFRelease(item)
        if secret is None:
            raise CredentialError("钥匙串没有此邮箱的专用密码，请先运行 setup。")
        return secret

    def set(self, account, secret):
        previous, item = self._find(account)
        del previous
        raw, name = secret.encode("utf-8"), account.encode("utf-8")
        try:
            if item:
                status = self.sec.SecKeychainItemModifyAttributesAndData(item, None, len(raw), raw)
            else:
                item = C.c_void_p()
                status = self.sec.SecKeychainAddGenericPassword(
                    None, len(SERVICE), SERVICE, len(name), name,
                    len(raw), raw, C.byref(item),
                )
            if status:
                raise CredentialError("无法保存邮箱专用密码，请解锁本机钥匙串后重试。")
        finally:
            if item:
                self.cf.CFRelease(item)
