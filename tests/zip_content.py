"""Zip packages compared by what they hold, not by how deflate packed it.

CPython 3.14's Windows builds compress with zlib-ng instead of zlib (as do the Linux
distributions that ship zlib-ng as their zlib): the same entries deflate to different,
equally valid bytes.  A golden that is a hash of a package's bytes then holds only where
deflate is zlib's own.  :func:`content_sha256` hashes every entry's name, metadata and
uncompressed bytes, in order -- everything the package says -- and :func:`stock_deflate`
tells whether this Python's deflate is zlib's, where the bytes are compared as well.
"""

from __future__ import annotations

import hashlib
import io
import zlib
import zipfile

# A few kilobytes of slide-like XML and what zlib (1.2.x and 1.3.x) deflates it to, as
# zipfile does (default level, raw stream).  zlib-ng compresses it differently.
_PROBE = b"".join(b"<a:p><a:r><a:t>%d %s</a:t></a:r></a:p>" % (i * i % 997, b"slide"[: i % 5])
                  for i in range(4000))
_ZLIB = "629b530296b7d832986087efafc63c5b67d6941047f32888d6071c07a1e7b6e5"


def stock_deflate() -> bool:
    """Whether deflate here gives zlib's own bytes, so a package's bytes can be compared."""
    compressor = zlib.compressobj(-1, zlib.DEFLATED, -15)
    return hashlib.sha256(compressor.compress(_PROBE) + compressor.flush()).hexdigest() == _ZLIB


def zlib_ng() -> bool:
    """Whether this Python says its zlib is zlib-ng (3.14+ on Windows, some distributions)."""
    return hasattr(zlib, "ZLIBNG_VERSION") or "ng" in zlib.ZLIB_RUNTIME_VERSION.lower()


def is_zip(data: bytes) -> bool:
    return data.startswith(b"PK\x03\x04")


def content_sha256(data: bytes) -> str:
    """A zip package's entries -- names, metadata, uncompressed bytes, in order -- hashed.

    An entry that is a package itself (a chart's embedded workbook) is hashed the same way,
    since its bytes are deflate's too."""
    digest = hashlib.sha256()
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        digest.update(repr(archive.comment).encode())
        for info in archive.infolist():
            entry = archive.read(info)
            nested = is_zip(entry)
            digest.update(repr((info.filename, info.date_time, info.compress_type,
                                info.create_system, info.external_attr, info.flag_bits,
                                info.extra, info.comment,
                                None if nested else info.file_size)).encode())
            digest.update(content_sha256(entry).encode() if nested
                          else hashlib.sha256(entry).digest())
    return digest.hexdigest()
