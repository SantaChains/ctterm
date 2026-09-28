# -*- coding: utf-8 -*-
"""CE 表单块（Encoding="Ascii85"）解码 / 解析 / 回写。

编码链（依据 CE 源码行为与公开逆向资料实证）：
  blob 文本 --字符集映射--> RFC1924 Base85 --b85decode--> zlib raw 流
  --inflate--> [4 字节前缀] + TPF0 表单流

字符集：CE 用"XML 安全"的 85 字符集（不含 & ' " < >），
与 RFC1924（0-9A-Za-z!#$%&()*+-;<=>?@^_`{|}~）逐位映射。
这解释了为何标准 base64.a85decode 报 Non-Ascii85 digit（blob 含 v y { }）。

TPF0 流：Delphi/VCL 表单序列化。属性值类型字节：
  2=Byte 3=Word 6=String(pascal) 7=Enum 8=False 9=True 10=Binary
翻译对象：type 6 的字符串属性（Caption/Text/Hint 等）。
回写：仅替换目标字符串字节，其余原样；重压缩（zlib raw）后 CE 可正常加载。
"""
from __future__ import annotations

import re
import struct
import zlib
from dataclasses import dataclass, field

# CE 字符集 → RFC1924 Base85 字符集
_CE_DIGITS = ('0123456789'
              'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
              'abcdefghijklmnopqrstuvwxyz'
              '!#$%()*+,-./:;=?@[]^_{}')
_RFC1924 = ('0123456789'
            'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
            'abcdefghijklmnopqrstuvwxyz'
            '!#$%&()*+-;<=>?@^_`{|}~')
assert len(_CE_DIGITS) == 85 and len(_RFC1924) == 85
_MAP = str.maketrans(_CE_DIGITS, _RFC1924)               # 解码：str → str
_MAP_BYTES = bytes.maketrans(_RFC1924.encode("ascii"),
                             _CE_DIGITS.encode("ascii"))  # 编码：bytes → bytes


def decode_blob(blob: str) -> bytes:
    """Ascii85 blob → TPF0 表单流（含 4 字节前缀）。"""
    import base64
    raw = base64.b85decode(blob.strip().translate(_MAP))
    data = zlib.decompress(raw, -zlib.MAX_WBITS)
    return data[4:]          # 前 4 字节为长度前缀（dword LE），TPF0 流不需要


def encode_form(tpf0: bytes) -> str:
    """TPF0 表单流 → blob 文本（前缀=流长度的 dword LE；raw deflate，
    与 decode_blob 的 -MAX_WBITS 对应）。"""
    import base64
    payload = struct.pack('<I', len(tpf0)) + tpf0
    comp = zlib.compressobj(6, zlib.DEFLATED, -zlib.MAX_WBITS)
    raw = comp.compress(payload) + comp.flush()
    return base64.b85encode(raw).translate(_MAP_BYTES).decode("ascii")


# ---------------------------------------------------------------------------
# TPF0 解析 / 序列化（保留未知属性字节）

@dataclass
class Prop:
    name: str
    ptype: int            # 2/3/6/7/8/9/10
    value: bytes | int | None   # type6/7: bytes；2: int；3: int；8/9: None；10: raw
    raw: bytes = b''      # type10 等未知类型的原始字节（含类型字节）


@dataclass
class Obj:
    obj_type: str
    obj_name: str
    props: list = field(default_factory=list)      # Prop | Obj


def read_pstr(data: bytes, off: int) -> tuple[bytes, int]:
    n = data[off]
    return data[off + 1: off + 1 + n], off + 1 + n


def parse_form(data: bytes) -> Obj:
    assert data[:4] == b'TPF0', f'非 TPF0 表单流: {data[:8]!r}'
    obj, off = _parse_object(data, 4)
    return obj


def _parse_object(data: bytes, off: int) -> tuple[Obj, int]:
    """TPF0 对象：class pstr + name pstr + 属性表(空名终止) + 子对象表(0 终止)。
    实证：属性表终结与子表终结是两个独立 0 字节（BulletLogForm 布局）。"""
    cls, off = read_pstr(data, off)
    name, off = read_pstr(data, off)
    obj = Obj(cls.decode("utf-8", "replace"), name.decode("utf-8", "replace"))
    # ---- 属性表 ----
    while True:
        fstart = off
        fname, off2 = read_pstr(data, off)
        if not fname:
            off = off2            # 属性表终结（空属性名）
            break
        ptype = data[off2]
        vstart = off2 + 1
        vend = _skip_value(data, vstart, ptype)
        if ptype == 6:            # String：唯一解析类型（Ident 为枚举名，禁译）
            val, _ = read_pstr(data, vstart)
            obj.props.append(Prop(fname.decode("utf-8", "replace"), ptype, val))
        else:                     # 其余 raw 透传（属性名+类型字节+值）
            obj.props.append(Prop(fname.decode("utf-8", "replace"), ptype, None,
                                  data[fstart:vend]))
        off = vend
    # ---- 子对象表 ----
    while off < len(data) and data[off] != 0:
        child, off = _parse_object(data, off)
        obj.props.append(child)
    if off < len(data) and data[off] == 0:
        off += 1                  # 子表终结符
    return obj, off


def _skip_value(data: bytes, off: int, ptype: int) -> int:
    """按 TPF0 (Delphi/FPC TWriter) 语义跳过值字节，返回新 off。"""
    if ptype in (6, 7):        # String/Ident: pstr
        _, off = read_pstr(data, off)
    elif ptype == 2:           # Int8
        off += 1
    elif ptype == 3:           # Int16
        off += 2
    elif ptype == 4:           # Int32
        off += 4
    elif ptype == 5:           # Extended
        off += 10
    elif ptype in (8, 9, 13):  # False/True/Nil
        pass
    elif ptype == 10:          # Binary: dword len + data
        n, = struct.unpack_from("<I", data, off); off += 4 + n
    elif ptype == 11:          # Set: pstr 序列，空 pstr 终止
        while data[off]:
            _, off = read_pstr(data, off)
        off += 1
    elif ptype == 12:          # LString: dword len + data
        n, = struct.unpack_from("<I", data, off); off += 4 + n
    elif ptype == 1:           # List: 值序列，0 字节终止
        while data[off]:
            off = _skip_value(data, off + 1, data[off])
        off += 1
    elif ptype == 14:          # Collection: item 序列
        off = _skip_collection(data, off)
    elif ptype == 15:          # Single
        off += 4
    elif ptype == 16:          # Currency
        off += 8
    elif ptype == 17:          # Date
        off += 8
    elif ptype == 18:          # WString: dword 字符数 + UTF-16LE
        n, = struct.unpack_from("<I", data, off); off += 4 + n * 2
    elif ptype == 19:          # Int64
        off += 8
    elif ptype == 20:          # UTF8String: dword len + data
        n, = struct.unpack_from("<I", data, off); off += 4 + n
    elif ptype == 21:          # Double
        off += 8
    else:
        raise ValueError(f"TPF0 未支持类型 {ptype} @ off={off}")
    return off


def _skip_collection(data: bytes, off: int) -> int:
    """vaCollection：{0=结束 | 1=item 起始（其后为属性序列，空属性名终止）}。"""
    while True:
        b = data[off]
        if b == 0:
            return off + 1
        if b == 1:
            off += 1
            while True:
                fname, off2 = read_pstr(data, off)
                if not fname:
                    off = off2      # item 终止 = 空属性名（无额外 0，实证）
                    break
                off = _skip_value(data, off2 + 1, data[off2])
        else:
            raise ValueError(f"collection 内未知标记 {b} @ {off}")


def serialize_form(obj: Obj) -> bytes:
    out = bytearray(b"TPF0")
    _ser_object(obj, out)
    return bytes(out)


def _pstr(s: bytes) -> bytes:
    assert len(s) < 256, "pascal string > 255"
    return bytes([len(s)]) + s


def _ser_object(obj: Obj, out: bytearray):
    out += _pstr(obj.obj_type.encode("utf-8"))
    out += _pstr(obj.obj_name.encode("utf-8"))
    for p in obj.props:
        if isinstance(p, Obj):
            continue
        if p.ptype == 6 and isinstance(p.value, bytes) and not p.raw:
            out += _pstr(p.name.encode("utf-8"))
            out.append(p.ptype)
            out += _pstr(p.value)
        elif p.raw:
            out += p.raw          # 属性名+类型字节+值，字节级原样
                                   # （含 type7 Ident：禁译枚举名，解析时按 raw 透传）
        else:
            out += _pstr(p.name.encode("utf-8"))
            out.append(p.ptype)
            v = p.value if isinstance(p.value, bytes) else str(p.value).encode("utf-8")
            out += _pstr(v)
    out.append(0)                 # 属性表终结
    for p in obj.props:
        if isinstance(p, Obj):
            _ser_object(p, out)
    out.append(0)                 # 子对象表终结


def collect_strings(root: Obj) -> list[tuple[Obj, Prop]]:
    """收集全部 type6 字符串属性（含嵌套对象）。"""
    out = []
    def walk(o: Obj):
        for p in o.props:
            if isinstance(p, Obj):
                walk(p)
            elif p.ptype == 6 and isinstance(p.value, bytes):
                out.append((o, p))
    walk(root)
    return out


def apply_form_translations(root: Obj, lookup: dict) -> tuple[int, int]:
    """按 {en: zh} 命中替换表单内字符串，返回 (替换数, 超长跳过数)。

    键规范化：源中的 \\r\\n 统一为 \\n（xlsx 往返会把 \\r 转义成 _x000D_
    破坏哈希——实测 "Tips:..." 行）。译文长度守卫：pstr 单字节长度上限
    255，超长译文拒绝写入（防止破坏 TPF0 流）。跳过数显式返回，
    由调用方计入统计——静默丢弃会让"看似已译实则未译"无从排查。"""
    n = 0
    skipped_long = 0
    for o, p in collect_strings(root):
        try:
            s = p.value.decode("utf-8")
        except UnicodeDecodeError:
            continue
        key = s.replace("\r\n", "\n").replace("\r", "\n")
        dst = lookup.get(key)
        if not dst or dst == key:
            continue
        if len(dst.encode("utf-8")) < 255:
            p.value = dst.encode("utf-8")
            n += 1
        else:
            skipped_long += 1
    return n, skipped_long
