"""Shared path and input-format safety checks for docforge exporters.

The converters in :mod:`docforge` only support Word's OOXML document
package.  A file ending in ``.docx`` is not enough to establish that: Office
also uses ZIP packages for PowerPoint and Excel, and legacy Word files are
OLE2 containers.  Keep the probe here so every source-consuming command can
reject those inputs before it creates a work directory or starts Pandoc.
"""

from __future__ import annotations

from pathlib import Path
import zipfile


UNSUPPORTED_FORMAT = "UNSUPPORTED_FORMAT"
INVALID_FORMAT = "INVALID_FORMAT"

_OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_ZIP_MAGICS = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")
_WORD_EXTENSIONS = frozenset({".docx", ".docm", ".dotx", ".dotm"})
_UNSUPPORTED_EXTENSIONS = frozenset({
    ".doc", ".dot", ".ppt", ".pptx", ".pptm", ".pot", ".potx", ".potm", ".pps", ".ppsm",
    ".sldx", ".sldm", ".xls", ".xlsx", ".xlsm", ".xlt", ".xltx", ".xltm", ".xlam",
})
_PRESENTATION_EXTENSIONS = frozenset({".ppt", ".pptx", ".pptm", ".pot", ".potx", ".potm", ".pps", ".ppsm", ".sldx", ".sldm"})
_SPREADSHEET_EXTENSIONS = frozenset({".xls", ".xlsx", ".xlsm", ".xlt", ".xltx", ".xltm", ".xlam"})


class FormatError(ValueError):
    """Stable, user-facing input-format failure.

    ``code`` is deliberately separate from the message so callers can make
    an exit/status decision without parsing prose.  Messages never include a
    source path or a parser's raw exception text.
    """

    code: str = INVALID_FORMAT

    def __init__(self, message: str, *, detected: str | None = None) -> None:
        super().__init__(message)
        self.detected = detected


class UnsupportedFormatError(FormatError):
    code = UNSUPPORTED_FORMAT


class InvalidFormatError(FormatError):
    code = INVALID_FORMAT


# Short aliases keep the typed result convenient for callers that do not use
# the ``Error`` suffix while preserving the explicit class names above.
UnsupportedFormat = UnsupportedFormatError
InvalidFormat = InvalidFormatError


class InputFormat(str):
    """String-like format result with the normalized source extension."""

    def __new__(cls, name: str, extension: str) -> "InputFormat":
        value = str.__new__(cls, name)
        value.extension = extension
        return value

    @property
    def name(self) -> str:
        return str(self)

    @property
    def format(self) -> str:
        """Alias used by callers that prefer an explicit field name."""

        return str(self)


def _package_format(names: set[str]) -> str | None:
    """Identify an OOXML package from its part names.

    Looking at names avoids parsing arbitrary package XML and is sufficient
    for dispatch.  The main part names are stable across transitional and
    strict OOXML packages.
    """

    normalized = {name.lstrip("/") for name in names}
    if "word/document.xml" in normalized:
        return "docx"
    if "ppt/presentation.xml" in normalized:
        return "pptx"
    if "xl/workbook.xml" in normalized:
        return "xlsx"
    return None


def _detect_format(path: Path) -> InputFormat:
    """Probe *path* and enforce the Word DOCX input contract.

    The extension and package magic are checked together.  A valid Word
    package with a non-Word extension is ``INVALID_FORMAT``; PowerPoint,
    Excel, and OLE2 inputs are ``UNSUPPORTED_FORMAT`` even when renamed to
    ``.docx``.  Invalid ZIPs and unknown bytes are ``INVALID_FORMAT``.
    """

    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(path)
    extension = path.suffix.casefold()
    try:
        with path.open("rb") as stream:
            magic = stream.read(8)
    except OSError:
        # Preserve normal filesystem errors, while keeping format failures
        # deterministic and free of parser diagnostics.
        raise

    if magic.startswith(_OLE2_MAGIC):
        raise UnsupportedFormatError("OLE2 documents are not supported", detected="ole2")

    package: str | None = None
    if magic[:4] in _ZIP_MAGICS:
        try:
            with zipfile.ZipFile(path) as archive:
                package = _package_format(set(archive.namelist()))
        except (OSError, RuntimeError, zipfile.BadZipFile, zipfile.LargeZipFile):
            raise InvalidFormatError("input is not a valid OOXML package") from None
    elif extension in _UNSUPPORTED_EXTENSIONS:
        # The extension itself is enough to reject known non-Word formats when
        # a damaged file no longer has a readable magic header.
        detected = "pptx" if extension in _PRESENTATION_EXTENSIONS else "xlsx" if extension in _SPREADSHEET_EXTENSIONS else "ole2"
        raise UnsupportedFormatError(f"{detected.upper()} documents are not supported", detected=detected)
    else:
        raise InvalidFormatError("input is not a DOCX package")

    if package in {"pptx", "xlsx"}:
        raise UnsupportedFormatError(f"{package.upper()} documents are not supported", detected=package)
    if package is None and extension in _UNSUPPORTED_EXTENSIONS:
        detected = "pptx" if extension in _PRESENTATION_EXTENSIONS else "xlsx" if extension in _SPREADSHEET_EXTENSIONS else "ole2"
        raise UnsupportedFormatError(f"{detected.upper()} documents are not supported", detected=detected)
    if package != "docx":
        raise InvalidFormatError("input is not a DOCX package")
    if extension not in _WORD_EXTENSIONS:
        raise InvalidFormatError("DOCX package extension does not match", detected=package)
    return InputFormat("docx", extension)


def detect_format(path: Path) -> InputFormat:
    """Public alias for :func:`_detect_format`."""

    return _detect_format(path)


def validate_output_path(
    path: Path | None,
    *,
    label: str = "output",
    forbid_tokens: tuple[str, ...] = (),
) -> None:
    """Check a publish path.

    Core format handling does not apply manuscript naming rules. Pass
    ``forbid_tokens`` from an opt-in delivery policy to reject names that
    contain those tokens.
    """

    if path is None:
        return
    folded = path.name.casefold()
    for token in forbid_tokens:
        needle = token.casefold()
        if needle and needle in folded:
            raise ValueError(f"{label} path contains forbidden token {token!r}: {path.name}")


__all__ = [
    "INVALID_FORMAT",
    "UNSUPPORTED_FORMAT",
    "FormatError",
    "UnsupportedFormatError",
    "InvalidFormatError",
    "UnsupportedFormat",
    "InvalidFormat",
    "InputFormat",
    "_detect_format",
    "detect_format",
    "validate_output_path",
]
