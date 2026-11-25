#!/usr/bin/env python3
"""Build ReviewTracker.dotm (a Word macro template with a ribbon tab) from source.

Word templates keep their macros in word/vbaProject.bin, an OLE compound file
in the MS-OVBA format. This script writes that file directly from
ReviewTracker.bas so the template can be rebuilt without Word. The project is
written with an empty performance cache (version 0xFFFF), which tells Word to
compile the macros from source the first time it opens the template.

    python3 word/build_dotm.py
"""
import io
import os
import random
import struct
import uuid
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
BAS = os.path.join(HERE, "ReviewTracker.bas")
RIBBON = os.path.join(HERE, "customUI14.xml")
OUT = os.path.join(HERE, "ReviewTracker.dotm")
CODEPAGE = 1252
PROJECT_NAME = "ReviewTracker"


# ── MS-OVBA compression (section 2.4.1) ──────────────────────────────────────

def _copy_token_help(offset_in_chunk):
    bit_count = max((offset_in_chunk - 1).bit_length(), 4)
    length_mask = 0xFFFF >> bit_count
    max_length = (0xFFFF >> bit_count) + 3
    return bit_count, length_mask, max_length


def _compress_chunk(data):
    out = bytearray()
    pos = 0
    n = len(data)
    recent = {}  # 3-byte sequence -> positions where it starts (most recent last)
    while pos < n:
        flag_index = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= n:
                break
            best_len, best_off = 0, 0
            if pos > 0:
                bit_count, _, max_length = _copy_token_help(pos)
                window = 1 << bit_count
                for cand in reversed(recent.get(data[pos:pos + 3], ())):
                    if pos - cand > window:
                        break
                    length = 0
                    while (length < max_length and pos + length < n
                           and data[cand + length] == data[pos + length]):
                        length += 1
                    if length > best_len:
                        best_len, best_off = length, pos - cand
                        if length == max_length:
                            break
            step = best_len if best_len >= 3 else 1
            if best_len >= 3:
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack("<H", token)
                flags |= 1 << bit
            else:
                out.append(data[pos])
            for p in range(pos, pos + step):
                recent.setdefault(data[p:p + 3], []).append(p)
            pos += step
        out[flag_index] = flags
    return bytes(out)


def compress(data):
    out = bytearray(b"\x01")
    for i in range(0, max(len(data), 1), 4096):
        chunk = data[i:i + 4096]
        comp = _compress_chunk(chunk)
        if len(comp) < 4096:
            header = 0xB000 | (len(comp) + 2 - 3)
            out += struct.pack("<H", header) + comp
        else:  # store raw
            header = 0x3000 | (4096 + 2 - 3)
            out += struct.pack("<H", header) + chunk.ljust(4096, b"\x00")
    return bytes(out)


# ── MS-OVBA data encryption for CMG / DPB / GC (section 2.4.3) ──────────────

def encrypt(data, project_id):
    seed = random.randint(0, 255)
    version = 2
    proj_key = sum(project_id.encode("ascii")) & 0xFF
    version_enc = seed ^ version
    proj_key_enc = seed ^ proj_key
    out = bytearray([seed, version_enc, proj_key_enc])
    unenc1, enc1, enc2 = proj_key, proj_key_enc, version_enc

    def push(b):
        nonlocal unenc1, enc1, enc2
        e = ((enc2 + unenc1) & 0xFF) ^ b
        out.append(e)
        enc2, enc1, unenc1 = enc1, e, b

    for _ in range((seed & 6) // 2):
        push(random.randint(0, 255))
    for b in struct.pack("<I", len(data)):
        push(b)
    for b in data:
        push(b)
    return out.hex().upper()


def decrypt(hexstr):
    """Inverse of encrypt(), used by the self-check."""
    data = bytes.fromhex(hexstr)
    seed, version_enc, proj_key_enc = data[0], data[1], data[2]
    proj_key = seed ^ proj_key_enc
    unenc1, enc1, enc2 = proj_key, proj_key_enc, version_enc
    pos = 3
    out = []

    def pull():
        nonlocal unenc1, enc1, enc2, pos
        e = data[pos]
        pos += 1
        b = e ^ ((enc2 + unenc1) & 0xFF)
        enc2, enc1, unenc1 = enc1, e, b
        return b

    for _ in range((seed & 6) // 2):
        pull()
    length = struct.unpack("<I", bytes(pull() for _ in range(4)))[0]
    return bytes(pull() for _ in range(length)), proj_key


# ── VBA project streams ─────────────────────────────────────────────────────

THIS_DOCUMENT = (
    'Attribute VB_Name = "ThisDocument"\r\n'
    'Attribute VB_Base = "1Normal.ThisDocument"\r\n'
    'Attribute VB_GlobalNameSpace = False\r\n'
    'Attribute VB_Creatable = False\r\n'
    'Attribute VB_PredeclaredId = True\r\n'
    'Attribute VB_Exposed = True\r\n'
    'Attribute VB_TemplateDerived = True\r\n'
    'Attribute VB_Customizable = True\r\n'
)


def rec(rid, payload):
    return struct.pack("<HI", rid, len(payload)) + payload


def dir_stream(modules):
    enc = lambda s: s.encode("cp1252")
    u16 = lambda s: s.encode("utf-16-le")
    d = bytearray()
    d += rec(0x0001, struct.pack("<I", 1))            # SYSKIND: 32-bit Windows (Word on any platform accepts it)
    d += rec(0x004A, struct.pack("<I", 2))            # COMPATVERSION
    d += rec(0x0002, struct.pack("<I", 0x0409))       # LCID
    d += rec(0x0014, struct.pack("<I", 0x0409))       # LCIDINVOKE
    d += rec(0x0003, struct.pack("<H", CODEPAGE))     # CODEPAGE
    d += rec(0x0004, enc(PROJECT_NAME))               # NAME
    d += rec(0x0005, b"") + rec(0x0040, b"")          # DOCSTRING + unicode
    d += rec(0x0006, b"") + rec(0x003D, b"")          # HELPFILEPATH 1 + 2
    d += rec(0x0007, struct.pack("<I", 0))            # HELPCONTEXT
    d += rec(0x0008, struct.pack("<I", 0))            # LIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)        # VERSION (major, minor)
    d += rec(0x000C, b"") + rec(0x003C, b"")          # CONSTANTS + unicode
    # Reference: OLE Automation (stdole)
    libid = "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation"
    d += rec(0x0016, enc("stdole")) + rec(0x003E, u16("stdole"))
    reg = struct.pack("<I", len(libid)) + enc(libid) + struct.pack("<IH", 0, 0)
    d += rec(0x000D, reg)
    # Modules
    d += rec(0x000F, struct.pack("<H", len(modules)))
    d += rec(0x0013, struct.pack("<H", 0xFFFF))       # PROJECTCOOKIE
    for name, kind, _ in modules:
        d += rec(0x0019, enc(name)) + rec(0x0047, u16(name))
        d += rec(0x001A, enc(name)) + rec(0x0032, u16(name))
        d += rec(0x001C, b"") + rec(0x0048, b"")
        d += rec(0x0031, struct.pack("<I", 0))        # text offset: source starts at 0
        d += rec(0x001E, struct.pack("<I", 0))        # help context
        d += rec(0x002C, struct.pack("<H", 0xFFFF))   # cookie
        d += struct.pack("<HI", 0x0021 if kind == "std" else 0x0022, 0)
        d += struct.pack("<HI", 0x002B, 0)            # module terminator
    d += struct.pack("<HI", 0x0010, 0)                # dir terminator
    return bytes(d)


def project_stream(modules, project_id):
    lines = ['ID="%s"' % project_id]
    for name, kind, _ in modules:
        lines.append(("Document=%s/&H00000000" % name) if kind == "doc" else "Module=%s" % name)
    lines += [
        'Name="%s"' % PROJECT_NAME,
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="%s"' % encrypt(struct.pack("<I", 0), project_id),
        'DPB="%s"' % encrypt(b"\x00", project_id),
        'GC="%s"' % encrypt(b"\xFF", project_id),
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    for name, _, _ in modules:
        lines.append("%s=0, 0, 0, 0, C" % name)
    return ("\r\n".join(lines) + "\r\n").encode("cp1252")


def projectwm_stream(modules):
    out = bytearray()
    for name, _, _ in modules:
        out += name.encode("cp1252") + b"\x00" + name.encode("utf-16-le") + b"\x00\x00"
    return bytes(out + b"\x00\x00")


# ── minimal OLE compound file writer (version 3, 512-byte sectors) ──────────

FREESECT, ENDOFCHAIN, FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


def cfb(tree):
    """tree: {"name": bytes | dict}. Returns the compound file bytes."""
    entries = []  # dicts: name, type (1 storage, 2 stream, 5 root), data, children

    def add(name, node, etype):
        e = {"name": name, "type": etype, "data": node if etype == 2 else b"", "kids": []}
        entries.append(e)
        index = len(entries) - 1
        if etype != 2:
            for k, v in node.items():
                e["kids"].append(add(k, v, 1 if isinstance(v, dict) else 2))
        return index

    add("Root Entry", tree, 5)

    # Mini stream holds streams under 4096 bytes in 64-byte mini sectors.
    mini = bytearray()
    minifat = []
    for e in entries:
        if e["type"] == 2 and len(e["data"]) < 4096:
            data = e["data"]
            if not data:
                e["start"], e["size"] = ENDOFCHAIN, 0
                continue
            first = len(mini) // 64
            nsec = (len(data) + 63) // 64
            mini += data.ljust(nsec * 64, b"\x00")
            minifat += [first + i + 1 for i in range(nsec - 1)] + [ENDOFCHAIN]
            e["start"], e["size"] = first, len(data)

    sectors = []   # list of 512-byte blocks (after the header)
    fat = []

    def alloc(data):
        if not data:
            return ENDOFCHAIN
        first = len(sectors)
        n = (len(data) + 511) // 512
        for i in range(n):
            sectors.append(data[i * 512:(i + 1) * 512].ljust(512, b"\x00"))
            fat.append(first + i + 1 if i < n - 1 else ENDOFCHAIN)
        return first

    for e in entries:
        if e["type"] == 2 and len(e["data"]) >= 4096:
            e["start"], e["size"] = alloc(e["data"]), len(e["data"])
    root = entries[0]
    root["start"], root["size"] = alloc(bytes(mini)), len(mini)
    minifat_bytes = b"".join(struct.pack("<I", x) for x in minifat)
    minifat_start = alloc(minifat_bytes) if minifat else ENDOFCHAIN
    n_minifat = (len(minifat_bytes) + 511) // 512

    # Directory: give each storage's children a simple red-black tree
    # (all black, sorted, as a right-leaning chain - valid per spec).
    for e in entries:
        e["left"] = e["right"] = e["child"] = NOSTREAM
        e.setdefault("start", 0)
        e.setdefault("size", 0)

    def key(i):
        n = entries[i]["name"]
        return (len(n), n.upper())

    for e in entries:
        kids = sorted(e["kids"], key=key)
        if kids:
            e["child"] = kids[0]
            for a, b in zip(kids, kids[1:]):
                entries[a]["right"] = b

    dir_bytes = bytearray()
    for e in entries:
        name = e["name"].encode("utf-16-le") + b"\x00\x00"
        d = name.ljust(64, b"\x00")
        d += struct.pack("<HBB", len(name), e["type"], 1)  # 1 = black
        d += struct.pack("<III", e["left"], e["right"], e["child"])
        d += b"\x00" * 16 + struct.pack("<I", 0) + b"\x00" * 16
        d += struct.pack("<IQ", e["start"], e["size"])
        dir_bytes += d
    empty = (b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0)
             + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM)
             + b"\x00" * 36 + struct.pack("<IQ", 0, 0))
    while len(dir_bytes) % 512:
        dir_bytes += empty
    dir_start = alloc(bytes(dir_bytes))

    # FAT sectors (may need to account for themselves)
    n_fat = 1
    while True:
        total = len(sectors) + n_fat
        if (total + 127) // 128 <= n_fat:
            break
        n_fat += 1
    fat_start = len(sectors)
    for i in range(n_fat):
        fat.append(FATSECT)
    fat += [FREESECT] * (n_fat * 128 - len(fat))
    for i in range(n_fat):
        sectors.append(b"".join(struct.pack("<I", x) for x in fat[i * 128:(i + 1) * 128]))
    assert n_fat <= 109

    header = bytearray()
    header += bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 16
    header += struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6)
    header += b"\x00" * 6
    header += struct.pack("<I", 0)                 # number of directory sectors (v3: 0)
    header += struct.pack("<I", n_fat)
    header += struct.pack("<I", dir_start)
    header += struct.pack("<I", 0)                 # transaction signature
    header += struct.pack("<I", 4096)              # mini stream cutoff
    header += struct.pack("<II", minifat_start, n_minifat)
    header += struct.pack("<II", ENDOFCHAIN, 0)    # no DIFAT sectors
    difat = [fat_start + i for i in range(n_fat)] + [FREESECT] * (109 - n_fat)
    header += b"".join(struct.pack("<I", x) for x in difat)
    assert len(header) == 512
    return bytes(header) + b"".join(sectors)


def vba_project_bin(source):
    project_id = "{%s}" % str(uuid.uuid4()).upper()
    modules = [("ThisDocument", "doc", THIS_DOCUMENT), ("ReviewTracker", "std", source)]
    vba = {
        "_VBA_PROJECT": struct.pack("<HHBH", 0x61CC, 0xFFFF, 0x00, 0x0000),
        "dir": compress(dir_stream(modules)),
    }
    for name, _, text in modules:
        vba[name] = compress(text.encode("cp1252"))
    tree = {
        "VBA": vba,
        "PROJECT": project_stream(modules, project_id),
        "PROJECTwm": projectwm_stream(modules),
    }
    return cfb(tree)


# ── the .dotm package ───────────────────────────────────────────────────────

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.ms-word.template.macroEnabledTemplate.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>
</Relationships>
"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>
"""

DOCUMENT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Review Tracker add-in template. Put this file in Word's Startup folder; you don't need to open it.</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>
"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>
"""

CORE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Review Tracker</dc:title>
</cp:coreProperties>
"""


def main():
    source = open(BAS, encoding="ascii", newline="").read()
    source = source.replace("\r\n", "\n").replace("\n", "\r\n")
    vba = vba_project_bin(source)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", DOCUMENT)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/vbaProject.bin", vba)
        z.writestr("docProps/core.xml", CORE)
        z.write(RIBBON, "customUI/customUI14.xml")
    print("Built", OUT)


if __name__ == "__main__":
    main()
