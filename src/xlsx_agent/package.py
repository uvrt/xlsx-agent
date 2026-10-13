"""Opening a workbook's package: what the bytes are, the limits, and the kind.

Everything that comes in is untrusted (ROADMAP.md, "Design rules", 6), so the bytes are
looked at before any of them is unpacked:

1. A compound file ([MS-CFB]) is not a package.  It is an encrypted workbook when it holds
   an ``EncryptedPackage`` stream ([MS-OFFCRYPTO] 2.3.4.4: how Office stores a package it
   encrypts), a ``.xls`` when it holds a ``Workbook`` or ``Book`` stream ([MS-XLS] 2.1.2:
   BIFF8's and BIFF5's), and something else otherwise.  A BIFF file without the compound
   file around it (BIFF2 to BIFF4, which begin with their BOF record, [MS-XLS] 2.4.21) is
   a ``.xls`` too.  Each is refused with its own code.
2. A ZIP goes through ooxml-edit's :func:`~ooxml_edit.tools.limits.check_package` (size,
   part count, expansion ratio, all from the central directory) before ooxml-edit's
   :class:`~ooxml_edit.opc.OpcPackage` reads it.
3. The package's main part must declare one of ``ooxml_common.kinds``' Excel kinds
   (``xlsx``, ``xlsm``, ``xltx``, ``xltm``).  A ``.xlsb``'s main part declares the binary
   workbook's type ([MS-XLSB] 2.1.1) and is refused; any other package is not a workbook.

Every refusal is an :class:`OpenError` whose :attr:`~OpenError.code` says why, and whose
message names no content of the file.
"""

from __future__ import annotations

import io
import os
import struct
import warnings
import zipfile

from ooxml_common.kinds import KINDS, kind_mismatch, kind_of
from ooxml_edit.opc import REL_OFFICE_DOCUMENT, OpcPackage
from ooxml_edit.tools.limits import LimitError, Limits, check_package

#: The package-level relationship to the main part in a strict package (ISO/IEC 29500
#: strict; ``REL_OFFICE_DOCUMENT`` is the transitional one).
REL_OFFICE_DOCUMENT_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"

#: A ``.xlsb``'s main part ([MS-XLSB] 2.1.1).
XLSB_CONTENT_TYPE = "application/vnd.ms-excel.sheet.binary.macroEnabled.main"

#: Every reason a workbook is refused, by :attr:`OpenError.code`.
REFUSALS: dict[str, str] = {
    "xls": "a binary Excel 97-2003 workbook (.xls); save it as .xlsx",
    "xlsb": "a binary Excel workbook (.xlsb); save it as .xlsx",
    "encrypted": "an encrypted (password-protected) workbook; remove the password and save it again",
    "not_a_workbook": "not an Excel workbook",
    "damaged": "a damaged file that cannot be read as a package",
    "limit": "over a size, part-count or expansion limit",
}

#: [MS-CFB] 2.2: the compound file header's signature.
CFB_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
#: [MS-XLS] 2.4.21: the BOF record's type for BIFF2, BIFF3 and BIFF4, little-endian, which
#: such a file begins with (BIFF5 and BIFF8 are always in a compound file).
BIFF_BOF = (b"\x09\x00", b"\x09\x02", b"\x09\x04")

_FREE = 0xFFFFFFFF
_END_OF_CHAIN = 0xFFFFFFFE


class OpenError(ValueError):
    """A file refused by :meth:`~xlsx_agent.Workbook.open`.  :attr:`code` is one of
    :data:`REFUSALS`' keys; the message says what the file is and names none of its
    content."""

    def __init__(self, code: str, detail: str | None = None) -> None:
        if code not in REFUSALS:
            raise ValueError(f"not a refusal code: {code!r}")
        self.code = code
        message = REFUSALS[code]
        super().__init__(f"{message}: {detail}" if detail else message)


class KindMismatchWarning(UserWarning):
    """A workbook opened from a path whose extension names another kind than the one its
    main part declares (a ``.xlsx`` that is a template, say).  It is read all the same;
    :attr:`~xlsx_agent.Workbook.kind` is what the package declares."""


def read_source(source, limits: Limits) -> bytes:
    """The bytes of a path, bytes-like object or binary file; a file over
    ``limits.max_package_bytes`` is refused before it is read."""
    if isinstance(source, (bytes, bytearray, memoryview)):
        return bytes(source)
    if isinstance(source, (str, os.PathLike)):
        size = os.stat(source).st_size
        if size > limits.max_package_bytes:
            raise OpenError("limit", f"{size} bytes; at most {limits.max_package_bytes}")
        with open(source, "rb") as handle:
            return handle.read()
    if hasattr(source, "read"):
        data = source.read(limits.max_package_bytes + 1)
        if not isinstance(data, bytes):
            raise TypeError("a file opened in binary mode is needed")
        if len(data) > limits.max_package_bytes:
            raise OpenError("limit", f"over {limits.max_package_bytes} bytes")
        return data
    raise TypeError(f"cannot open a workbook from {type(source).__name__}: give bytes, a path or an OpcPackage")


def open_package(data: bytes, limits: Limits) -> OpcPackage:
    """Check ``data`` as the module's notes say and open it as a package; :class:`OpenError`
    when it is refused."""
    if data.startswith(CFB_SIGNATURE):
        raise OpenError(*_compound_file_refusal(data))
    if data[:2] in BIFF_BOF:
        raise OpenError("xls", "a BIFF file")
    if not data.startswith(b"PK"):
        raise OpenError("not_a_workbook", "neither a ZIP package nor a compound file")
    if not zipfile.is_zipfile(io.BytesIO(data)):
        raise OpenError("damaged", "its ZIP directory cannot be read")
    try:
        check_package(data, limits, "the workbook")
    except LimitError as exc:
        raise OpenError("limit", str(exc)) from None
    try:
        return OpcPackage.open(data)
    except (zipfile.BadZipFile, zipfile.LargeZipFile, ValueError, EOFError, OSError) as exc:
        raise OpenError("damaged", type(exc).__name__) from None


def workbook_part(package: OpcPackage, source=None) -> tuple[str, str]:
    """The main part and its kind (``"xlsx"``...), checked: :class:`OpenError` unless the
    main part declares an Excel kind.  ``source``, when it is a path, has its extension
    compared with that kind, and a :class:`KindMismatchWarning` is given when they differ."""
    part = None
    for rel in package.relationships("").values():
        if rel.type in (REL_OFFICE_DOCUMENT, REL_OFFICE_DOCUMENT_STRICT) and rel.target_part is not None:
            part = rel.target_part
            break
    if part is None or not package.has_part(part):
        raise OpenError("not_a_workbook", "the package has no main part")
    content_type = package.content_type(part)
    if content_type is not None and content_type.split(";")[0].strip() == XLSB_CONTENT_TYPE:
        raise OpenError("xlsb")
    kind = kind_of(content_type)
    if kind is None or KINDS[kind].application != "excel":
        what = f"a .{kind} package" if kind is not None else "a package of no Excel kind"
        raise OpenError("not_a_workbook", what)
    if isinstance(source, (str, os.PathLike)):
        mismatch = kind_mismatch(content_type, source)
        if mismatch is not None:
            warnings.warn(KindMismatchWarning(str(mismatch)), stacklevel=3)
    return part, kind


# -- compound files ----------------------------------------------------------------------------


def _compound_file_refusal(data: bytes) -> tuple[str, str]:
    """``(code, detail)`` for a compound file, from the names in its directory."""
    try:
        names = compound_file_names(data)
    except ValueError as exc:
        return "damaged", f"a compound file whose directory cannot be read ({exc})"
    if "EncryptedPackage" in names:
        return "encrypted", "an encrypted package in a compound file"
    if "Workbook" in names or "Book" in names:
        return "xls", "a BIFF workbook in a compound file"
    return "not_a_workbook", "a compound file with no workbook stream"


def compound_file_names(data: bytes) -> set[str]:
    """The names of the storages and streams in a compound file's directory ([MS-CFB] 2.2,
    2.3, 2.6), read without reading any stream.  ``ValueError`` when the header or the
    sector chains do not hold together."""
    if len(data) < 512 or not data.startswith(CFB_SIGNATURE):
        raise ValueError("no compound file header")
    shift = struct.unpack_from("<H", data, 0x1E)[0]
    if shift not in (9, 12):
        raise ValueError("an unknown sector size")
    size = 1 << shift
    sectors = max((len(data) - size) // size, 0)

    def sector(number: int) -> bytes:
        if number >= sectors:
            raise ValueError("a sector past the end of the file")
        start = (number + 1) * size
        return data[start:start + size]

    # The FAT's own sectors: the header's 109 DIFAT entries, then the DIFAT chain.
    fat_count, first_directory = struct.unpack_from("<II", data, 0x2C)
    next_difat, difat_count = struct.unpack_from("<II", data, 0x44)
    fat_sectors = [n for n in struct.unpack_from("<109I", data, 0x4C) if n != _FREE]
    per_sector = size // 4
    for _ in range(min(difat_count, sectors)):
        if next_difat in (_END_OF_CHAIN, _FREE):
            break
        entries = struct.unpack(f"<{per_sector}I", sector(next_difat))
        fat_sectors.extend(n for n in entries[:-1] if n != _FREE)
        next_difat = entries[-1]
    fat_sectors = fat_sectors[:fat_count] if fat_count else fat_sectors
    fat: list[int] = []
    for number in fat_sectors:
        fat.extend(struct.unpack(f"<{per_sector}I", sector(number)))

    names: set[str] = set()
    number, seen = first_directory, 0
    while number != _END_OF_CHAIN:
        seen += 1
        if seen > sectors or number >= len(fat):
            raise ValueError("a broken directory chain")
        block = sector(number)
        for offset in range(0, size, 128):
            length, kind = struct.unpack_from("<HB", block, offset + 0x40)
            if kind in (1, 2, 5) and 2 <= length <= 64:
                names.add(block[offset:offset + length - 2].decode("utf-16-le", "replace"))
        number = fat[number]
    return names
