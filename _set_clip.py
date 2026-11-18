"""把粘贴包 3 行 12 列按 Excel 剪贴板约定写进剪贴板（CF_UNICODETEXT）。
约定：列用 \t，行用 \r\n；含 \t/\r/\n/\" 的单元格整体加双引号，内部双引号翻倍。
"""
import ctypes
import ctypes.wintypes as W
import sys

import openpyxl

GMEM_MOVEABLE = 0x0002
CF_UNICODETEXT = 13

k32 = ctypes.windll.kernel32
u32 = ctypes.windll.user32
k32.GlobalAlloc.restype = ctypes.c_void_p
k32.GlobalAlloc.argtypes = [W.UINT, ctypes.c_size_t]
k32.GlobalLock.restype = ctypes.c_void_p
k32.GlobalLock.argtypes = [ctypes.c_void_p]
k32.GlobalUnlock.argtypes = [ctypes.c_void_p]
u32.OpenClipboard.argtypes = [ctypes.c_void_p]
u32.SetClipboardData.argtypes = [W.UINT, ctypes.c_void_p]
u32.SetClipboardData.restype = ctypes.c_void_p

P = sys.argv[1] if len(sys.argv) > 1 else r"D:/tk/paste_pack.xlsx"   # 默认粘贴包路径，建议用参数传
wb = openpyxl.load_workbook(P)
ws = wb["批量粘贴"]
hdr = [str(c.value) for c in ws[1]]
# 空单元格必须是真正的空字符串：之前填了「（空）」字面量，客户端会当成封面名称去校验并报「请输入正确的封面名称」
rows = [["" if c.value is None else str(c.value) for c in r] for r in ws.iter_rows(min_row=2)]
print("列顺序:", " | ".join(hdr))


def cell(v):
    if any(ch in v for ch in "\t\r\n\""):
        return '"' + v.replace('"', '""') + '"'
    return v


text = "\r\n".join("\t".join(cell(c) for c in r) for r in rows)
print("行数=", len(rows), "字符=", len(text))

data = text.encode("utf-16-le") + b"\x00\x00"
if not u32.OpenClipboard(None):
    sys.exit("OpenClipboard 失败")
try:
    u32.EmptyClipboard()
    h = k32.GlobalAlloc(GMEM_MOVEABLE, len(data))
    if not h:
        sys.exit("GlobalAlloc 失败")
    p = k32.GlobalLock(h)
    ctypes.memmove(p, data, len(data))
    k32.GlobalUnlock(h)
    if not u32.SetClipboardData(CF_UNICODETEXT, h):
        sys.exit("SetClipboardData 失败")
finally:
    u32.CloseClipboard()
print("[剪贴板] 已写入 %d 行 x %d 列" % (len(rows), len(hdr)))
