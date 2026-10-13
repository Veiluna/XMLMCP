"""Publish a complete report directory atomically without replacing an existing path."""
from __future__ import annotations

import ctypes
import errno
import os
import sys
from pathlib import Path

from .errors import XMLMCPError, require


def publish_bundle(source: Path, destination: Path) -> None:
    try:
        if os.name == "nt":
            # Windows rename rejects even an existing empty destination directory.
            source.rename(destination)
            return
        libc = ctypes.CDLL(None, use_errno=True)
        if sys.platform == "darwin":
            rename = libc.renamex_np
            rename.argtypes = [ctypes.c_char_p, ctypes.c_char_p, ctypes.c_uint]
            rename.restype = ctypes.c_int
            # Apple sys/stdio.h: RENAME_EXCL = 0x00000004.
            result = rename(os.fsencode(source), os.fsencode(destination), 4)
        elif sys.platform.startswith("linux"):
            rename = libc.renameat2
            rename.argtypes = [ctypes.c_int, ctypes.c_char_p, ctypes.c_int, ctypes.c_char_p, ctypes.c_uint]
            rename.restype = ctypes.c_int
            # AT_FDCWD and RENAME_NOREPLACE from the Linux / libc headers.
            result = rename(-100, os.fsencode(source), -100, os.fsencode(destination), 1)
        else:
            raise XMLMCPError("REPORT_PUBLICATION_UNSUPPORTED", "当前平台不支持报告目录的原子且不覆盖发布")
        if result:
            code = ctypes.get_errno()
            raise OSError(code, os.strerror(code), str(destination))
    except AttributeError as exc:
        raise XMLMCPError("REPORT_PUBLICATION_UNSUPPORTED", "系统缺少不覆盖目录发布接口") from exc
    except OSError as exc:
        require(exc.errno not in (errno.EEXIST, errno.ENOTEMPTY) and not destination.exists(),
                "OUTPUT_EXISTS", "输出目录已存在，未覆盖现有内容")
        raise XMLMCPError("REPORT_PUBLICATION_FAILED", "报告目录发布失败，未交付结果", error=str(exc)) from exc
