"""Resolve the visible run formatting of WordprocessingML text.

Word records many run properties that do not change the rendered page:
proofing language, complex-script twins of bold/italic/size, font hints for
characters the run does not contain, or an explicit black color on a black
style.  :class:`StyleSheet` resolves docDefaults, table, paragraph and
character styles plus direct formatting, and :meth:`StyleSheet.visible`
returns only what changes the glyphs of the given text.
"""

from __future__ import annotations

from typing import Iterable

from lxml import etree

from .ooxml import W, Token, as_dict, flatten

_TOGGLES = (
    "b", "i", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss",
    "imprint", "vanish",
)
_VALUED = ("vertAlign", "position", "spacing", "w", "kern", "sz", "em", "effect", "rtl", "cs")
_FALSE = {"0", "false", "off", "none"}
_SLOTS = ("ascii", "hAnsi", "eastAsia", "cs")
_IGNORED_PREFIXES = (
    "lang", "noProof", "webHidden", "specVanish", "bCs", "iCs", "szCs", "oMath",
    "snapToGrid",
)


def _is_cjk(char: str) -> bool:
    code = ord(char)
    return (
        0x2E80 <= code <= 0x9FFF
        or 0xF900 <= code <= 0xFAFF
        or 0xFE30 <= code <= 0xFE4F
        or 0xFF00 <= code <= 0xFFEF
        or 0xAC00 <= code <= 0xD7AF
        or 0x20000 <= code <= 0x2FA1F
    )


def _is_complex(char: str) -> bool:
    code = ord(char)
    return 0x0590 <= code <= 0x08FF or 0x0E00 <= code <= 0x0E7F


def font_slots(text: str, hint: str | None) -> set[str]:
    """Font slots Word uses for the characters of ``text``.

    Word always draws ASCII with the ``ascii`` font; an ``eastAsia`` hint only
    moves the remaining shared characters (dashes, quotes, symbols) to the
    East Asian font.  LibreOffice applies the hint more widely, so its layout
    is not the reference for round-trip checks.
    """
    slots: set[str] = set()
    for char in text:
        if ord(char) < 0x80:
            slots.add("ascii")
        elif _is_cjk(char):
            slots.add("eastAsia")
        elif _is_complex(char):
            slots.add("cs")
        elif hint == "eastAsia":
            slots.add("eastAsia")
        else:
            slots.add("hAnsi")
    return slots


class StyleSheet:
    """Style inheritance from a package's ``word/styles.xml``."""

    def __init__(self, styles_root: etree._Element | None) -> None:
        self.styles: dict[str, etree._Element] = {}
        self.names: dict[str, str] = {}
        self.default_paragraph: str | None = None
        self.default_character: str | None = None
        self.defaults: list[Token] = []
        if styles_root is None:
            return
        rpr = styles_root.find(f"./{{{W}}}docDefaults/{{{W}}}rPrDefault/{{{W}}}rPr")
        self.defaults = flatten(rpr)
        for style in styles_root.findall(f"./{{{W}}}style"):
            style_id = style.get(f"{{{W}}}styleId")
            if not style_id:
                continue
            self.styles[style_id] = style
            name = style.find(f"./{{{W}}}name")
            self.names[style_id] = name.get(f"{{{W}}}val") if name is not None else style_id
            if style.get(f"{{{W}}}default") in {"1", "true"}:
                kind = style.get(f"{{{W}}}type")
                if kind == "paragraph":
                    self.default_paragraph = style_id
                elif kind == "character":
                    self.default_character = style_id

    def chain(self, style_id: str | None) -> list[etree._Element]:
        result: list[etree._Element] = []
        seen: set[str] = set()
        while style_id and style_id in self.styles and style_id not in seen:
            seen.add(style_id)
            style = self.styles[style_id]
            result.append(style)
            based = style.find(f"./{{{W}}}basedOn")
            style_id = based.get(f"{{{W}}}val") if based is not None else None
        return list(reversed(result))

    def style_tokens(self, style_id: str | None, part: str) -> list[list[Token]]:
        return [flatten(style.find(f"./{{{W}}}{part}")) for style in self.chain(style_id)]

    def outline_level(self, style_id: str | None) -> int | None:
        level = None
        for style in self.chain(style_id):
            node = style.find(f"./{{{W}}}pPr/{{{W}}}outlineLvl")
            if node is not None:
                level = int(node.get(f"{{{W}}}val", "9"))
        return level

    def layers(
        self,
        paragraph_style: str | None,
        run_tokens: Iterable[Token],
        table_style: str | None = None,
    ) -> list[list[Token]]:
        run_tokens = list(run_tokens)
        layers = [self.defaults]
        if table_style:
            layers += self.style_tokens(table_style, "rPr")
        layers += self.style_tokens(paragraph_style or self.default_paragraph, "rPr")
        run_style = as_dict(run_tokens).get("rStyle@val")
        if run_style:
            layers += self.style_tokens(run_style, "rPr")
        layers.append(run_tokens)
        return layers

    def visible(
        self,
        paragraph_style: str | None,
        run_tokens: Iterable[Token],
        text: str,
        table_style: str | None = None,
    ) -> tuple:
        """A hashable description of how ``text`` renders with these properties."""
        layers = self.layers(paragraph_style, run_tokens, table_style)
        merged: dict[str, str | None] = {}
        fonts: dict[str, str] = {}
        for layer in layers:
            layer_map = as_dict(layer)
            for slot in _SLOTS:
                theme = layer_map.get(f"rFonts@{slot}Theme")
                name = layer_map.get(f"rFonts@{slot}")
                if theme:
                    fonts[slot] = "theme:" + theme
                elif name:
                    fonts[slot] = name
            if "rFonts@hint" in layer_map:
                fonts["hint"] = layer_map["rFonts@hint"] or ""
            for attribute in ("val", "eastAsia", "bidi"):
                if f"lang@{attribute}" in layer_map:
                    fonts[f"lang:{attribute}"] = layer_map[f"lang@{attribute}"] or ""
            for key, value in layer_map.items():
                if key.startswith("rFonts") or key.startswith("lang"):
                    continue
                top = key.split("@", 1)[0].split(".", 1)[0]
                # A layer that sets an element replaces the inherited one.
                for existing in [item for item in merged if item.split("@", 1)[0].split(".", 1)[0] == top]:
                    if existing not in layer_map:
                        del merged[existing]
                merged[key] = value
        signature: list[tuple] = []
        for name in _TOGGLES:
            present = name in merged or f"{name}@val" in merged
            if present and (merged.get(f"{name}@val") or "1").lower() not in _FALSE:
                signature.append((name, True))
        underline = merged.get("u@val")
        if "u" in merged and underline is None:
            underline = "single"
        if underline and underline != "none":
            signature.append(("u", underline, merged.get("u@color")))
        highlight = merged.get("highlight@val")
        if highlight and highlight != "none":
            signature.append(("highlight", highlight))
        color = merged.get("color@val")
        theme = merged.get("color@themeColor")
        if theme:
            signature.append(("color", theme, merged.get("color@themeTint"), merged.get("color@themeShade")))
        elif color and color.lower() not in {"auto", "000000"}:
            signature.append(("color", color.upper()))
        fill = merged.get("shd@fill")
        if fill and fill.lower() != "auto" and merged.get("shd@val") != "nil":
            signature.append(("shd", fill.upper(), merged.get("shd@val")))
        for name in _VALUED:
            value = merged.get(f"{name}@val")
            if name in {"rtl", "cs"}:
                if name in merged or value:
                    signature.append((name, (value or "1").lower() not in _FALSE))
                continue
            if value is not None and not (name == "vertAlign" and value == "baseline"):
                if name in {"position", "spacing", "kern"} and value == "0":
                    continue
                if name == "w" and value == "100":
                    continue
                signature.append((name, value))
        for key, value in merged.items():
            top = key.split("@", 1)[0].split(".", 1)[0]
            if top in {"bdr", "fitText", "eastAsianLayout", "rPrChange", "ins", "del", "moveFrom", "moveTo"}:
                signature.append((key, value))
        slots = font_slots(text, fonts.get("hint"))
        for slot in sorted(slots):
            if slot == "hAnsi":
                signature.append(("font", slot, fonts.get("hAnsi") or fonts.get("ascii")))
            else:
                signature.append(("font", slot, fonts.get(slot)))
            # The East Asian and complex-script languages select the face of
            # theme fonts and the line-breaking rules for those characters.
            if slot == "eastAsia":
                signature.append(("lang", slot, fonts.get("lang:eastAsia")))
            elif slot == "cs":
                signature.append(("lang", slot, fonts.get("lang:bidi")))
        return tuple(signature)


def is_ignored_key(key: str) -> bool:
    top = key.lstrip("-").split("@", 1)[0].split(".", 1)[0]
    return top in _IGNORED_PREFIXES


__all__ = ["StyleSheet", "font_slots", "is_ignored_key"]
