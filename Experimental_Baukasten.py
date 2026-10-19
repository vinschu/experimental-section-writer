# -*- coding: utf-8 -*-
"""
author: vinschu
Modified: preview shows procedure at startup and updates immediately when a procedure is selected.
Keeps previous features (dark theme, drag & drop, undo, button color customizations, etc.)
"""
import sys
import json
import re
import time
import math
from copy import deepcopy
from pathlib import Path
import os

os.environ.setdefault(
    "QT_QPA_FONTDIR",
    str(Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"),
)

from PySide6 import QtWidgets, QtGui, QtCore
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt

import xml.etree.ElementTree as ET
import html as html_std

# Standard-Datei-Pfade
if getattr(sys, 'frozen', False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).parent.resolve()

PROCEDURES_FILE = APP_DIR / "procedures.json"
SOLVENTS_FILE = APP_DIR / "solvents.json"
TRIVIALS_FILE = APP_DIR / "trivial_names.json"
REACTION_LEARNING_FILE = APP_DIR / "reaction_learning.json"

# Dark stylesheet applied at application level, inspired by modern Apple software.
DARK_STYLESHEET = """
    QWidget { background: #1c1c1e; color: #f5f5f7; font-family: "SF Pro Display", "Segoe UI", sans-serif; font-size: 13px; }
    QMainWindow, QDialog { background: #1c1c1e; }
    QFrame { background: #242426; color: #f5f5f7; border: 1px solid #3a3a3c; border-radius: 8px; }
    QMenuBar { background: #242426; color: #f5f5f7; border-bottom: 1px solid #3a3a3c; padding: 2px 6px; }
    QMenuBar::item { background: transparent; color: #f5f5f7; padding: 6px 9px; border-radius: 5px; }
    QMenuBar::item:selected { background: #3a3a3c; }
    QMenu { background: #2c2c2e; color: #f5f5f7; border: 1px solid #48484a; padding: 5px; }
    QMenu::item { padding: 7px 24px 7px 10px; border-radius: 5px; }
    QMenu::item:selected { background: #3a3a3c; }

    QListWidget, QTextEdit, QLineEdit, QPlainTextEdit, QComboBox, QSpinBox, QDoubleSpinBox {
        background: #2c2c2e; color: #f5f5f7; border: 1px solid #48484a; border-radius: 7px;
        padding: 6px 8px; selection-background-color: #0a84ff; selection-color: #ffffff;
    }
    QListWidget::item { padding: 5px 7px; border-radius: 5px; }
    QListWidget::item:selected { background: #164a78; color: #ffffff; }
    QLineEdit:focus, QPlainTextEdit:focus, QTextEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus {
        border: 2px solid #0a84ff; padding: 5px 7px;
    }
    QComboBox { padding-right: 30px; }
    QComboBox::drop-down { subcontrol-origin: padding; subcontrol-position: top right; width: 28px; border-left: 1px solid #48484a; }
    QComboBox::down-arrow { width: 0; height: 0; border-left: 5px solid transparent; border-right: 5px solid transparent; border-top: 6px solid #a1a1a6; }
    QComboBox:hover { border-color: #636366; }
    QComboBox QAbstractItemView { background: #2c2c2e; color: #f5f5f7; border: 1px solid #48484a; padding: 4px; outline: 0; }
    QComboBox QAbstractItemView::item { min-height: 20px; padding: 2px 8px; border-radius: 5px; }
    QComboBox QAbstractItemView::item:hover, QComboBox QAbstractItemView::item:selected { background: #3a3a3c; color: #ffffff; }

    QPushButton {
        background-color: #3a3a3c; color: #f5f5f7; padding: 7px 13px; border-radius: 7px;
        border: 1px solid #48484a; font-weight: 600;
    }
    QPushButton:hover { background-color: #48484a; }
    QPushButton:pressed { background-color: #636366; }
    QPushButton:disabled { background-color: #2c2c2e; color: #636366; border-color: #3a3a3c; }

    QPushButton[secondary="true"] { background-color: #3a3a3c; color: #d1d1d6; border: 1px solid #48484a; }
    QPushButton[secondary="true"]:hover { background-color: #48484a; }
    QPushButton[secondary="true"]:pressed { background-color: #636366; }

    QPushButton#copyPreviewBtn { background-color: #0a84ff; color: #ffffff; border: 1px solid #0a84ff; padding: 8px 16px; }
    QPushButton#copyPreviewBtn:hover { background-color: #409cff; border-color: #409cff; }
    QPushButton#copyPreviewBtn:pressed { background-color: #0066cc; }

    QPushButton#undoBtn { background-color: #3a2427; color: #ff6961; border: 1px solid #713238; }
    QPushButton#undoBtn:hover { background-color: #51282d; }
    QPushButton#undoBtn:pressed { background-color: #713238; }

    QPushButton#editOutputBtn { background-color: #243b2c; color: #63d471; border: 1px solid #356340; }
    QPushButton#editOutputBtn:hover { background-color: #2c5137; }
    QPushButton#editOutputBtn:pressed { background-color: #356340; }

    QToolButton#formatToolButton {
        background-color: #3a3a3c; color: #f5f5f7; border: 1px solid #48484a;
        border-radius: 4px; padding: 2px 5px; min-width: 22px; max-width: 28px;
        font-weight: 600;
    }
    QToolButton#formatToolButton:hover { background-color: #48484a; }
    QToolButton#formatToolButton:checked { background-color: #0a84ff; border-color: #0a84ff; }

    QLabel { color: #f5f5f7; }
    QStatusBar { background: #242426; color: #a1a1a6; border-top: 1px solid #3a3a3c; }
    QToolTip { background: #f5f5f7; color: #1c1c1e; border: 0; padding: 5px 7px; }
    QScrollBar:vertical { background: transparent; width: 12px; margin: 2px; }
    QScrollBar::handle:vertical { background: #636366; min-height: 28px; border-radius: 6px; }
    QScrollBar::handle:vertical:hover { background: #8e8e93; }
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
"""

def load_procedures():
    if not PROCEDURES_FILE.exists():
        sample = [
            {
                "id": "proc1",
                "name": "General esterification",
                "description": "Fischer esterification variant",
                "template": "To a stirred solution of {acid} (1.0 equiv) in {solvent} (10 mL) was added {alcohol} (1.2 equiv). The mixture was cooled to {temperature} and {catalyst} (5 mol%) was added. The reaction was stirred for {time} and then quenched with {quench}."
            }
        ]
        PROCEDURES_FILE.write_text(json.dumps(sample, indent=2, ensure_ascii=False))
    try:
        with open(PROCEDURES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return []
    procs = []
    if isinstance(data, dict):
        if "procedures" in data and isinstance(data["procedures"], list):
            raw_list = data["procedures"]
            procs = [p for p in raw_list if isinstance(p, dict)]
        else:
            items = [(k, v) for k, v in data.items() if isinstance(v, dict)]
            if items:
                for k, v in items:
                    copy = dict(v)
                    if "id" not in copy or not copy.get("id"):
                        copy["id"] = k
                    procs.append(copy)
            else:
                if all(isinstance(v, (str, int, float, bool, list, dict, type(None))) for v in data.values()):
                    procs = [data]
    elif isinstance(data, list):
        procs = [p for p in data if isinstance(p, dict)]
    else:
        return []
    normalized = []
    for i, p in enumerate(procs):
        pid = str(p.get("id") or p.get("ID") or f"proc{i+1}")
        name = p.get("name") or p.get("title") or ""
        desc = p.get("description") or ""
        template = p.get("template") or p.get("body") or ""
        normalized.append({
            "id": pid,
            "name": str(name),
            "description": str(desc),
            "template": str(template)
        })
    try:
        raw = PROCEDURES_FILE.read_text(encoding="utf-8")
        try:
            raw_obj = json.loads(raw)
        except Exception:
            raw_obj = None
        if raw_obj != normalized:
            PROCEDURES_FILE.write_text(json.dumps(normalized, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass
    return normalized

def save_procedures(procs):
    with open(PROCEDURES_FILE, "w", encoding="utf-8") as f:
        json.dump(procs, f, indent=2, ensure_ascii=False)

def load_solvents():
    if not SOLVENTS_FILE.exists():
        default = ["THF", "Et2O", "Hexane", "DCM", "MeOH"]
        SOLVENTS_FILE.write_text(json.dumps(default, indent=2))
    try:
        with open(SOLVENTS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, list):
                return [str(x) for x in data]
    except Exception:
        pass
    return []

def save_solvents(solvents):
    with open(SOLVENTS_FILE, "w", encoding="utf-8") as f:
        json.dump(solvents, f, indent=2, ensure_ascii=False)

def load_trivial_names():
    if not TRIVIALS_FILE.exists():
        TRIVIALS_FILE.write_text(json.dumps({}, indent=2))
    try:
        with open(TRIVIALS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                normalized = {}
                for k, v in data.items():
                    if not isinstance(k, str):
                        continue
                    kk = remove_subscript_unicode(strip_markup_to_plain_ascii(k)).strip()
                    if kk:
                        normalized[kk] = str(v)
                return normalized
    except Exception:
        pass
    return {}

def save_trivial_names(mapping):
    try:
        with open(TRIVIALS_FILE, "w", encoding="utf-8") as f:
            json.dump(mapping, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

def load_reaction_learning():
    if not REACTION_LEARNING_FILE.exists():
        return {}
    try:
        with open(REACTION_LEARNING_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}

def save_reaction_learning(mapping):
    try:
        with open(REACTION_LEARNING_FILE, "w", encoding="utf-8") as f:
            json.dump(mapping, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

def chemical_signature(chemicals):
    formulas = {
        remove_subscript_unicode(str(chemical.get("formula", ""))).replace(" ", "").lower()
        for chemical in chemicals
        if chemical.get("formula")
    }
    return "|".join(sorted(formulas))

def extract_placeholders(template_text):
    return list(dict.fromkeys(re.findall(r"\{([a-zA-Z0-9_]+)\}", template_text)))

def suggest_reaction_index(chemicals, procedures, reaction_learning=None):
    """Return the safest matching procedure index, or None when uncertain."""
    if not chemicals or not procedures:
        return None
    reaction_learning = reaction_learning or {}
    signature = chemical_signature(chemicals)
    learned = reaction_learning.get(signature, {})
    formula_source = " ".join(chemical.get("formula", "") for chemical in chemicals)
    source = " ".join(
        f"{chemical.get('formula', '')} {chemical.get('display', '')}".lower()
        for chemical in chemicals
    )
    has_halogen = bool(re.search(r"Br|Cl|I", formula_source)) or bool(re.search(r"iod|halogen", source))
    has_boron = bool(re.search(r"B(?=[A-Z]|$)|bor|borester|boronic", formula_source + " " + source))
    has_lithium = bool(re.search(r"Li", formula_source)) or bool(re.search(r"buli|n-buli", source))
    has_nitrogen = bool(re.search(r"N(?=[A-Z]|$)", formula_source)) or bool(re.search(r"amine|amino", source))
    has_cobalt = bool(re.search(r"Co", formula_source))
    scores = []
    for index, procedure in enumerate(procedures):
        name = procedure.get("name", "").lower()
        procedure_id = str(procedure.get("id", ""))
        score = 0
        if "suzuki" in name and has_halogen and has_boron and not has_lithium:
            score = 8
        elif "miyaura" in name and has_halogen and has_boron:
            score = 7
        elif "borylation" in name and has_lithium and has_boron:
            score = 9
        elif "alkylation" in name and has_halogen and has_nitrogen:
            score = 8
        elif "trimerization" in name and len(chemicals) >= 3 and has_cobalt:
            score = 8
        score += min(int(learned.get(procedure_id, 0)), 3)
        scores.append((score, index))
    best_score, best_index = max(scores, default=(0, None))
    return best_index if best_score >= 7 else None

def _chemical_text(chemical):
    return f"{chemical.get('formula', '')} {chemical.get('display', '')}".lower()

def _chemical_role_score(role, chemical, procedure_name):
    """Score a chemical for a placeholder; zero means no safe match."""
    text = _chemical_text(chemical)
    formula = remove_subscript_unicode(str(chemical.get("formula", "")))
    role = role.lower()
    procedure_name = procedure_name.lower()
    is_halide = bool(re.search(r"br|cl|iod|brom|chlor", text, re.IGNORECASE))
    is_boron = bool(re.search(r"bor|borester|boronic", text)) or bool(re.search(r"B(?=[A-Z]|$)", formula))
    is_lithium = bool(re.search(r"buli|n-buli|lithium", text)) or bool(re.search(r"Li", formula))
    is_nitrogen = bool(re.search(r"amine|amino|nitrogen", text)) or bool(re.search(r"N(?=[A-Z]|$)", formula))
    is_catalyst = bool(re.search(r"catalyst|palladium|nickel|cobalt|pd|ni|co|cu", text))
    is_base = bool(re.search(r"base|carbonate|hydroxide|phosphate|amine|triethylamine|tBuONa|tBuOK", text))

    if role in {"c1", "c2"}:
        if "suzuki" in procedure_name or "miyaura" in procedure_name:
            wanted = is_halide if role == "c1" else is_boron
            return 10 if wanted else 0
        if "borylation" in procedure_name:
            return 9 if role == "c1" and not is_boron and not is_lithium else 0
        if "alkylation" in procedure_name:
            wanted = is_nitrogen if role == "c1" else is_halide
            return 8 if wanted else 0
    elif role == "buli":
        return 12 if is_lithium else 0
    elif role == "borester":
        return 12 if is_boron else 0
    elif role == "catalyst":
        return 8 if is_catalyst else 0
    elif role == "base":
        return 7 if is_base else 0
    return 0

def suggest_chemical_mapping(chemicals, placeholders, procedure, reaction_learning=None):
    """Return only unambiguous automatic assignments for chemical placeholders."""
    if not chemicals or not procedure:
        return {}
    procedure_name = str(procedure.get("name", ""))
    signature = chemical_signature(chemicals)
    learned = (reaction_learning or {}).get(signature, {})
    procedure_id = str(procedure.get("id", ""))
    learned_mapping = {}
    if isinstance(learned, dict):
        mappings_by_procedure = learned.get("__mapping_by_procedure__", {})
        if isinstance(mappings_by_procedure, dict):
            learned_mapping = mappings_by_procedure.get(procedure_id, {}) or {}
        if not learned_mapping:
            legacy_mapping = learned.get("__mapping__", {})
            if isinstance(legacy_mapping, dict) and all(isinstance(value, str) for value in legacy_mapping.values()):
                learned_mapping = legacy_mapping
    available = list(chemicals)
    product_reserved = any(placeholder.lower() == "product" for placeholder in placeholders)
    if product_reserved and available:
        available = available[:-1]
    result = {}

    for placeholder in placeholders:
        if "solvent" in placeholder.lower():
            continue
        if placeholder.lower() == "product" and chemicals:
            last_chemical = chemicals[-1]
            result[placeholder] = last_chemical.get("display") or last_chemical.get("formula")
            continue
        learned_value = learned_mapping.get(placeholder)
        if learned_value:
            learned_match = next(
                (c for c in available if c.get("display") == learned_value or c.get("formula") == learned_value),
                None,
            )
            if learned_match is not None:
                result[placeholder] = learned_match.get("display") or learned_match.get("formula")
                available.remove(learned_match)
                continue

        scored = sorted(
            ((_chemical_role_score(placeholder, chemical, procedure_name), chemical) for chemical in available),
            key=lambda item: item[0],
            reverse=True,
        )
        if not scored or scored[0][0] < 7:
            continue
        best_score, best = scored[0]
        second_score = scored[1][0] if len(scored) > 1 else 0
        if best_score == second_score:
            continue
        result[placeholder] = best.get("display") or best.get("formula")
        available.remove(best)
    return result

# ---------------------------
# Helpers: subscript mapping and others
# ---------------------------

SUBSCRIPT_TRANSLATION = str.maketrans({
    "0": "₀",
    "1": "₁",
    "2": "₂",
    "3": "₃",
    "4": "₄",
    "5": "₅",
    "6": "₆",
    "7": "₇",
    "8": "₈",
    "9": "₉"
})
_UNICODE_SUB_TO_ASCII = {v: k for k, v in {
    "0": "₀",
    "1": "₁",
    "2": "₂",
    "3": "₃",
    "4": "₄",
    "5": "₅",
    "6": "₆",
    "7": "₇",
    "8": "₈",
    "9": "₉"
}.items()}
SUBSCRIPT_REVERSE_TRANSLATION = str.maketrans(_UNICODE_SUB_TO_ASCII)
UNICODE_SUBSCRIPT_DIGITS = set(_UNICODE_SUB_TO_ASCII.keys())

def subscript_digits(s: str) -> str:
    if not s:
        return s
    return s.translate(SUBSCRIPT_TRANSLATION)

def remove_subscript_unicode(s: str) -> str:
    if not s:
        return s
    return s.translate(SUBSCRIPT_REVERSE_TRANSLATION)

def normalize_commas_in_number_string(s: str) -> str:
    if not s:
        return s
    return re.sub(r'(?<=\d),(?=\d)', '.', s)

def is_zero_ml(s):
    if not s:
        return True
    raw = s.strip().lower().replace("\u00A0", " ").strip()
    if "ml" in raw:
        num = re.sub(r"[^\d\.,\-eE]", "", raw)
        num = num.replace(",", ".")
        try:
            val = float(num)
            return abs(val) < 1e-9
        except Exception:
            return False
    return False

def append_eq_if_present(s):
    if not s:
        return s
    if re.search(r"\beq\.?$", s, re.IGNORECASE):
        return s
    return f"{s} eq."

def apply_subscript_markup_legacy(s: str) -> str:
    if not s:
        return s
    s = re.sub(r'(?i)<sub>\s*([0-9]+)\s*</sub>',
               lambda m: m.group(1).translate(SUBSCRIPT_TRANSLATION),
               s)
    s = re.sub(r'_\{\s*([0-9]+)\s*\}',
               lambda m: m.group(1).translate(SUBSCRIPT_TRANSLATION),
               s)
    s = re.sub(r'_([0-9]+)',
               lambda m: m.group(1).translate(SUBSCRIPT_TRANSLATION),
               s)
    return s

_MARKUP_RE = re.compile(
    r'(?P<htmlsub><sub>.*?</sub>)'
    r'|_\{\s*([^}]+)\s*\}'
    r'|_([A-Za-z0-9])'
    r'|/\{\s*([^}]+)\s*\}'
    r'|/([A-Za-z0-9])',
    flags=re.IGNORECASE | re.DOTALL
)

def render_markup_to_html(s: str) -> str:
    if s is None:
        return ""
    bold_matches = list(re.finditer(r"(?is)<b>(.*?)</b>", s))
    if bold_matches:
        out = []
        last = 0
        for match in bold_matches:
            out.append(render_markup_to_html(s[last:match.start()]))
            out.append(f"<b>{render_markup_to_html(match.group(1))}</b>")
            last = match.end()
        out.append(render_markup_to_html(s[last:]))
        return "".join(out)
    out = []
    last = 0
    for m in _MARKUP_RE.finditer(s):
        start, end = m.span()
        if start > last:
            out.append(html_std.escape(s[last:start]))
        htmlsub = m.group(1)
        brace_sub = m.group(2)
        single_sub = m.group(3)
        brace_i = m.group(4)
        single_i = m.group(5)
        if htmlsub:
            inner = re.sub(r'(?i)^<sub>\s*', '', htmlsub)
            inner = re.sub(r'(?i)\s*</sub>$', '', inner)
            out.append(f"<sub>{html_std.escape(inner)}</sub>")
        elif brace_sub:
            out.append(f"<sub>{html_std.escape(brace_sub)}</sub>")
        elif single_sub:
            out.append(f"<sub>{html_std.escape(single_sub)}</sub>")
        elif brace_i:
            out.append(f"<i>{html_std.escape(brace_i)}</i>")
        elif single_i:
            out.append(f"<i>{html_std.escape(single_i)}</i>")
        else:
            out.append(html_std.escape(m.group(0)))
        last = end
    if last < len(s):
        out.append(html_std.escape(s[last:]))
    return "".join(out)

def strip_markup_to_plain_ascii(s: str) -> str:
    if s is None:
        return ""
    s = re.sub(r'(?i)<sub>\s*([^<]+?)\s*</sub>', r'\1', s)
    s = re.sub(r'(?i)<i>\s*([^<]+?)\s*</i>', r'\1', s)
    s = re.sub(r'(?i)</?b>', '', s)
    s = re.sub(r'_\{\s*([^}]+)\s*\}', r'\1', s)
    s = re.sub(r'_([A-Za-z0-9])', r'\1', s)
    s = re.sub(r'/\{\s*([^}]+)\s*\}', r'\1', s)
    s = re.sub(r'/([A-Za-z0-9])', r'\1', s)
    return s

# ---------------------------
# CDXML parsing and helpers (unchanged)
# ---------------------------

def local_name(el):
    tag = el.tag
    if isinstance(tag, str) and tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag

def extract_text_elem(elem):
    if elem is None:
        return ""
    return "".join(elem.itertext()).strip()

def _binary_cdx_semantic_values(formula, text_values):
    """Map CDX table tokens by units so omitted columns cannot shift values."""
    values = []
    for value in text_values:
        if (
            values
            and value in {"mg", "g", "mL", "M", "mmol", "mol", "%"}
            and re.fullmatch(r"\d+[,.]\d+", values[-1])
        ):
            values[-1] = f"{values[-1]}{value}"
        else:
            values.append(value)

    def number(value):
        return bool(re.fullmatch(r"\d+[,.]\d+", value))

    def has_unit(value, unit):
        return bool(re.fullmatch(rf"\d+[,.]\d+\s*{unit}", value, re.IGNORECASE))

    molecular_weight = next((value for value in values if number(value)), "")
    flags = [index for index, value in enumerate(values) if value in {"Yes", "No"}]
    flag_index = flags[0] if flags else -1
    equivalent = next((value for value in values[flag_index + 1:] if number(value)), "")
    masses = [value for value in values if has_unit(value, r"(?:mg|g)")]
    amounts = [value for value in values if has_unit(value, r"(?:mmol|mol)")]
    volumes = [value for value in values if has_unit(value, r"mL")]
    percents = [value for value in values if has_unit(value, r"%")]
    molarities = [value for value in values if has_unit(value, r"M")]
    densities = [value for value in values if has_unit(value, r"g/mL")]

    if not molecular_weight or not flag_index or not equivalent:
        return None
    result = [
        formula,
        molecular_weight,
        amounts[0] if amounts else "",
        values[flag_index],
        equivalent,
        masses[0] if masses else "",
        percents[0] if percents else "",
        molarities[0] if molarities else "",
        densities[0] if densities else "",
        volumes[0] if volumes else "",
        amounts[-1] if amounts else "",
        masses[-1] if masses else "",
    ]
    product_values = {
        "mass": masses[-1] if masses else "",
        "amount": amounts[-1] if amounts else "",
        "yield": percents[-1] if percents else "",
    }
    return result, product_values

def parse_binary_cdx_table(cdx_path):
    headers = [
        "Formula", "MW", "Limit Moles", "Limiting?", "Equivalents",
        "Sample Mass", "%Weight", "Molarity", "Density", "Volume",
        "Reactant Moles", "Reactant Mass",
    ]
    data = Path(cdx_path).read_bytes()
    formula_matches = [
        match for match in re.finditer(
            rb"(?<![A-Za-z])(?:[A-Z][a-z]?\d*)+(?=\x00)", data,
        )
        if any(byte >= ord("0") and byte <= ord("9") for byte in match.group())
    ]
    if not formula_matches:
        return []

    components = []
    for index, formula_match in enumerate(formula_matches):
        formula = formula_match.group().decode("ascii", errors="ignore")
        if formula == "CCs2O3":
            formula = "Cs2CO3"
        start = formula_match.end()
        end = formula_matches[index + 1].start() if index + 1 < len(formula_matches) else len(data)
        text_values = []
        for text_match in re.finditer(rb"[ -~]{2,}", data[start:end]):
            value = text_match.group().decode("latin1").strip()
            if value:
                text_values.append(value)

        combined_values = []
        unit_only = {"mg", "g", "mL", "M", "mmol", "mol", "%"}
        for value in text_values:
            if (
                combined_values
                and value in unit_only
                and re.fullmatch(r"\d+[,.]\d+", combined_values[-1])
            ):
                combined_values[-1] = f"{combined_values[-1]}{value}"
            else:
                combined_values.append(value)
        text_values = combined_values

        semantic_result = _binary_cdx_semantic_values(formula, text_values)
        if semantic_result is not None:
            semantic_values, product_values = semantic_result
            components.append({
                "formula": formula,
                "headers": headers,
                "values": semantic_values,
                "raw_fields": dict(zip(headers, semantic_values)),
                "product_values": product_values,
            })
            continue

        values = [formula]
        for value in text_values:
            if len(values) == 1 and re.fullmatch(r"\d+[,.]\d+", value):
                values.append(value)
            elif len(values) == 2 and re.search(r"(?:mmol|mol)$", value, re.IGNORECASE):
                values.append(value)
            elif len(values) == 3 and value in {"Yes", "No"}:
                values.append(value)
            elif len(values) == 3 and re.fullmatch(r"\d+[,.]\d+", value):
                values.append(value)
            elif len(values) >= 4 and len(values) < len(headers):
                if value == "mL" and values[-1].endswith("g"):
                    values[-1] += "/mL"
                elif (
                    re.fullmatch(r"\d+[,.]\d+\s*(?:mg|g|mL|M|mmol|mol|%|g/mL)", value, re.IGNORECASE)
                    or (len(values) == 4 and re.fullmatch(r"\d+[,.]\d+", value))
                    or value in {"Yes", "No"}
                ):
                    values.append(value)

        if len(values) >= 3:
            values.extend([""] * (len(headers) - len(values)))
            components.append({
                "formula": formula,
                "headers": headers,
                "values": values[:len(headers)],
                "raw_fields": dict(zip(headers, values[:len(headers)])),
            })
    return components

def parse_stoichiometry_components(cdxml_path):
    if Path(cdxml_path).suffix.lower() == ".cdx":
        return parse_binary_cdx_table(cdxml_path)
    try:
        tree = ET.parse(cdxml_path)
    except ET.ParseError as e:
        raise RuntimeError(f"XML parse error for {cdxml_path}: {e}")
    root = tree.getroot()
    stoich = None
    for el in root.iter():
        if isinstance(el.tag, str) and local_name(el).lower() == "stoichiometrygrid":
            stoich = el
            break
    if stoich is None:
        return []
    components = [c for c in stoich if local_name(c).lower() == "sgcomponent"]
    if not components:
        components = [el for el in stoich.iter() if local_name(el).lower() == "sgcomponent"]
    if not components:
        return []
    header_comp = next((c for c in components if (c.attrib.get("ComponentIsHeader") or "").lower() == "yes"), None)
    headers = []
    if header_comp is not None:
        sgdata_elems = [el for el in header_comp if local_name(el).lower() == "sgdatum"]
        for sg in sgdata_elems:
            t_elem = next((d for d in sg.iter() if local_name(d).lower() == "t"), None)
            headers.append(extract_text_elem(t_elem) or "")
    else:
        first = components[0] if components else None
        if first is not None:
            sgdata_elems = [el for el in first if local_name(el).lower() == "sgdatum"]
            for sg in sgdata_elems:
                t_elem = next((d for d in sg.iter() if local_name(d).lower() == "t"), None)
                headers.append(extract_text_elem(t_elem) or "")
    data_components = [c for c in components if c is not header_comp]
    if len(data_components) > 0:
        data_components = data_components[1:]
    parsed = []
    for comp in data_components:
        sgdatums = [el for el in comp if local_name(el).lower() == "sgdatum"]
        values = []
        for sg in sgdatums:
            t_elem = next((d for d in sg.iter() if local_name(d).lower() == "t"), None)
            values.append(extract_text_elem(t_elem) or "")
        if len(values) < len(headers):
            values += [""] * (len(headers) - len(values))
        raw_fields = {}
        for i, hdr in enumerate(headers):
            key = (hdr or "").strip() or f"col_{i}"
            raw_fields[key] = values[i] if i < len(values) else ""
        formula = ""
        for k in raw_fields.keys():
            if k.strip().lower() == "formula":
                formula = raw_fields[k].strip()
                break
        if not formula and values:
            formula = values[0].strip()
        formula = re.sub(r'^[\s\-_\.0-9:()]+', '', formula).strip()
        if formula:
            parsed.append({
                "formula": formula,
                "headers": headers,
                "values": values,
                "raw_fields": raw_fields
            })
    return parsed

def find_cdxml_files(root_folder):
    cdxml_files = []
    for dirpath, _, filenames in os.walk(root_folder):
        for fname in filenames:
            if fname.lower().endswith((".cdx", ".cdxml", ".xml")):
                full = os.path.join(dirpath, fname)
                try:
                    with open(full, "rb") as fh:
                        head = fh.read(400).lower()
                    if b"cdx" in head or b"cdxml" in head or b"chem" in head or fname.lower().endswith((".cdx", ".cdxml")):
                        cdxml_files.append(full)
                except Exception:
                    if fname.lower().endswith((".cdx", ".cdxml")):
                        cdxml_files.append(full)
    return cdxml_files

def format_component_display_by_index(component, is_last=False):
    """
    Extrahiert relevante Zeilen aus component['values'] und baut die Anzeige zusammen.

    Änderungen:
    - Für alle verwendeten Zeilen (row0, row4, row5, row9, row10, row11) wird zuerst
      normalize_commas_in_number_string angewendet und danach ein Leerzeichen zwischen
      Zahl und vielen gebräuchlichen Einheiten eingefügt (z. B. "10mg" -> "10 mg").
      Prozentangaben bleiben unverändert (z. B. "5%").
    - Verhalten für die letzte Zeile: Anzeige als "row1 (row9, row10, row11)" wie gewünscht.
    """
    import re

    vals = component.get("values", [])

    def get_last_col_value(row_index):
        try:
            if not vals:
                return ""
            first = vals[0]
            # Fall: values ist Liste von Zeilen (jede Zeile ist Liste/Tuple von Spalten)
            if isinstance(first, (list, tuple)):
                # Wenn row_index außerhalb liegt, fange Exception ab
                return (vals[row_index][-1] or "").strip() if row_index < len(vals) and vals[row_index] else ""
            # Fall: values ist Liste von Spalten (jede Spalte ist Liste/Tuple von Zeilen)
            last = vals[-1]
            if isinstance(last, (list, tuple)):
                return (last[row_index] or "").strip() if row_index < len(last) else ""
            # Fallback: flache Liste (einspaltig)
            return (vals[row_index] or "").strip() if row_index < len(vals) else ""
        except Exception:
            return ""

    def get_first_col_value(row_index):
        """Wird für formula/backoff verwendet (ähnlich zum Originalverhalten)."""
        try:
            if not vals:
                return ""
            first = vals[0]
            # Liste von Zeilen -> erste Spalte ist index 0
            if isinstance(first, (list, tuple)):
                return (vals[row_index][0] or "").strip() if row_index < len(vals) and vals[row_index] else ""
            # Liste von Spalten -> erste Spalte ist vals[0]
            if isinstance(vals[0], (list, tuple)):
                col0 = vals[0]
                return (col0[row_index] or "").strip() if row_index < len(col0) else ""
            # Fallback: flache Liste
            return (vals[row_index] or "").strip() if row_index < len(vals) else ""
        except Exception:
            return ""

    def space_before_units(s: str) -> str:
        """
        Fügt ein Leerzeichen zwischen Zahl und gebräuchlichen Einheiten ein.
        Belässt reine Prozentangaben '5%' unverändert.
        """
        if not s:
            return s
        # Entferne überflüssige Leerzeichen vor '%' (so bleibt '5%' ohne Leerzeichen)
        s = re.sub(r'\s+%', '%', s)

        # Erweiterbare Einheitensammlung (große Auswahl gebräuchlicher Einheiten)
        units = [
            r'mg', r'g', r'kg', r'µg', r'ug', r'μg',
            r'mol', r'mmol', r'µmol', r'umol', r'cmol', r'mM', r'M',
            r'mol/L', r'mol·L-1', r'mol L-1', r'mol·kg-1',
            r'L', r'l', r'mL', r'mmL', r'mml', r'ml', r'µL', r'uL', r'nL',
            r'µg/mL', r'ug/mL', r'wt%', r'v/v', r'v/v\%', r'ppm',
            r'eq\.?', r'equiv\.?', r'equiv', r'U', r'°C', r'K',
            r'bar', r'kPa', r'atm', r'mmHg', r'cm', r'mm', r'µm', r'nm'
        ]
        units_pattern = r'(?:' + r'|'.join(units) + r')\b'

        # Ersetze Muster wie "10mg", "1,0mol", "2e-3mol" -> "10 mg", "1,0 mol", "2e-3 mol"
        def repl(m):
            num = m.group('num')
            unit = m.group('unit')
            return f"{num} {unit}"

        s = re.sub(
            rf'(?i)(?P<num>[-+]?\d[\d\.,eE+\-]*)\s*(?P<unit>{units_pattern})',
            repl,
            s
        )

        return s

    def process_disp(raw: str) -> str:
        """Normalize commas and add spacing to units."""
        if not raw:
            return ""
        normalized = normalize_commas_in_number_string(raw)
        spaced = space_before_units(normalized)
        return spaced

    # Extract relevant raw values (using first/last-col heuristics)
    row0_raw = get_first_col_value(0)
    row4_raw = get_last_col_value(4)
    row5_raw = get_last_col_value(5)
    row9_raw = get_last_col_value(9)
    row10_raw = get_last_col_value(10)
    row11_raw = get_last_col_value(11)

    # Process displays for all these rows
    row0_disp = process_disp(row0_raw)
    row4_disp = process_disp(row4_raw)
    row5_disp = process_disp(row5_raw)
    row9_disp = process_disp(row9_raw)
    row10_disp = process_disp(row10_raw)
    row11_disp = process_disp(row11_raw)

    # row4_eq_disp should append "eq." if necessary (apply on processed form)
    row4_eq_disp = append_eq_if_present(row4_disp) if row4_disp else ""

    # chosen logic: prefer row9 unless it's an effective zero-ml, else row5
    chosen = row9_disp if (row9_raw and not is_zero_ml(row9_raw)) else row5_disp

    # formula fallback logic (retain original cleaning and subscript display)
    formula_raw = component.get("formula", "").strip() or row0_raw
    formula_raw = re.sub(r'^[\s\-_\.0-9:()]+', '', formula_raw).strip()
    formula_display = subscript_digits(formula_raw)

    if is_last:
        # For last row show "row1 (row9, row10, row11)" as requested.
        # For row1 we keep subscripted formula-like display (row0 may be a formula)
        row1_label = subscript_digits(row0_raw) if row0_raw else formula_display or formula_raw
        product_values = component.get("product_values", {})
        if product_values:
            row9_disp = process_disp(product_values.get("mass", ""))
            row10_disp = process_disp(product_values.get("amount", ""))
            row11_disp = process_disp(product_values.get("yield", ""))
        bracket_parts = []
        for part in (row9_disp, row10_disp, row11_disp):
            if part:
                bracket_parts.append(part)
        bracket = ", ".join(bracket_parts)
        display = f"{row1_label} ({bracket})" if bracket else row1_label
        return {"formula": formula_raw, "display": display, "details": {"row1": row1_label, "row9": row9_disp, "row10": row10_disp, "row11": row11_disp}}

    # Non-last: original behaviour: formula_display with bracket built from chosen, row10, row4_eq
    bracket_parts = []
    for part in (chosen, row10_disp, row4_eq_disp):
        if part:
            bracket_parts.append(part)
    bracket = ", ".join(bracket_parts)
    display = f"{formula_display} ({bracket})" if bracket else formula_display
    return {"formula": formula_raw, "display": display, "details": {"chosen": chosen, "row10": row10_disp, "row4_eq": row4_eq_disp}}
# ---------------------------
# Draggable list widget: ensure plain-text mime during drag
# ---------------------------

class DraggableListWidget(QtWidgets.QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragEnabled(True)
        self.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)

    def startDrag(self, supportedActions):
        item = self.currentItem()
        if not item:
            return
        mime = QtCore.QMimeData()
        mime.setText(item.text())
        drag = QtGui.QDrag(self)
        drag.setMimeData(mime)
        pix = QtGui.QPixmap(200, 20)
        pix.fill(QtGui.QColor("#2a2a2a"))
        painter = QtGui.QPainter(pix)
        painter.setPen(QtGui.QColor("#ffffff"))
        painter.drawText(6, 14, item.text()[:60])
        painter.end()
        drag.setPixmap(pix)
        drag.exec(QtCore.Qt.CopyAction)

# ---------------------------
# Preview text widget (accepts drops from lists)
# ---------------------------

class PreviewText(QtWidgets.QTextEdit):
    dropped = QtCore.Signal(str, QtCore.QPoint)
    drag_hovered = QtCore.Signal(str, QtCore.QPoint)
    drag_left = QtCore.Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(False)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        md = event.mimeData()
        if md.hasText() or md.hasHtml():
            event.setDropAction(QtCore.Qt.CopyAction)
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        md = event.mimeData()
        if md.hasText() or md.hasHtml():
            event.setDropAction(QtCore.Qt.CopyAction)
            text = md.text() if md.hasText() else md.html()
            pos = event.position().toPoint() if hasattr(event, 'position') else event.pos()
            self.drag_hovered.emit(text, pos)
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.drag_left.emit()
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        md = event.mimeData()
        if not (md.hasText() or md.hasHtml()):
            event.ignore()
            return
        text = md.text() if md.hasText() else md.html()
        pos = event.position().toPoint() if hasattr(event, 'position') else event.pos()
        self.dropped.emit(text, pos)
        event.setDropAction(QtCore.Qt.CopyAction)
        event.accept()

# ---------------------------
# Remaining dialogs (inherit app stylesheet)
# ---------------------------

class SolventEditorDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, solvent=None):
        super().__init__(parent)
        self.setWindowTitle("Solvent Editor")
        self.resize(400, 120)
        layout = QtWidgets.QVBoxLayout(self)
        form = QtWidgets.QFormLayout()
        self.name_edit = QtWidgets.QLineEdit()
        form.addRow("Solvent name:", self.name_edit)
        layout.addLayout(form)
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)
        if solvent:
            self.name_edit.setText(solvent)
    def get_data(self):
        return self.name_edit.text().strip()

class SolventManagerDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, solvents=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Solvents")
        self.resize(500, 350)
        self.solvents = list(solvents or [])
        layout = QtWidgets.QVBoxLayout(self)
        self.list_widget = QtWidgets.QListWidget()
        layout.addWidget(self.list_widget)
        self.refresh_list()
        btn_layout = QtWidgets.QHBoxLayout()
        add_btn = QtWidgets.QPushButton("Add")
        edit_btn = QtWidgets.QPushButton("Edit")
        del_btn = QtWidgets.QPushButton("Delete")
        up_btn = QtWidgets.QPushButton("↑")
        down_btn = QtWidgets.QPushButton("↓")
        up_btn.setToolTip("Move selected solvent up")
        down_btn.setToolTip("Move selected solvent down")
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(del_btn)
        btn_layout.addWidget(up_btn)
        btn_layout.addWidget(down_btn)
        layout.addLayout(btn_layout)
        add_btn.clicked.connect(self.add_solvent)
        edit_btn.clicked.connect(self.edit_solvent)
        del_btn.clicked.connect(self.delete_solvent)
        up_btn.clicked.connect(self.move_up)
        down_btn.clicked.connect(self.move_down)
        close_btn = QtWidgets.QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
    def refresh_list(self):
        self.list_widget.clear()
        for s in self.solvents:
            self.list_widget.addItem(subscript_digits(s))
    def add_solvent(self):
        dlg = SolventEditorDialog(self)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            val = dlg.get_data()
            if val:
                self.solvents.append(val)
                self.refresh_list()
    def edit_solvent(self):
        item = self.list_widget.currentItem()
        if not item:
            return
        idx = self.list_widget.currentRow()
        val = self.solvents[idx] if 0 <= idx < len(self.solvents) else ""
        dlg = SolventEditorDialog(self, val)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            newv = dlg.get_data()
            if newv:
                self.solvents[idx] = newv
                self.refresh_list()
    def delete_solvent(self):
        item = self.list_widget.currentItem()
        if not item:
            return
        idx = self.list_widget.currentRow()
        if idx < 0 or idx >= len(self.solvents):
            return
        confirm = QtWidgets.QMessageBox.question(self, "Delete", f"Delete solvent '{self.solvents[idx]}'?")
        if confirm == QtWidgets.QMessageBox.StandardButton.Yes:
            del self.solvents[idx]
            self.refresh_list()

    def move_up(self):
        idx = self.list_widget.currentRow()
        if idx <= 0 or idx >= len(self.solvents):
            return
        self.solvents[idx - 1], self.solvents[idx] = self.solvents[idx], self.solvents[idx - 1]
        self.refresh_list()
        self.list_widget.setCurrentRow(idx - 1)

    def move_down(self):
        idx = self.list_widget.currentRow()
        if idx < 0 or idx >= len(self.solvents) - 1:
            return
        self.solvents[idx], self.solvents[idx + 1] = self.solvents[idx + 1], self.solvents[idx]
        self.refresh_list()
        self.list_widget.setCurrentRow(idx + 1)

class ProcedureEditorDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, procedure=None):
        super().__init__(parent)
        self.setWindowTitle("Procedure Editor")
        self.resize(700, 400)
        layout = QtWidgets.QVBoxLayout(self)
        form = QtWidgets.QFormLayout()
        self.name_edit = QtWidgets.QLineEdit()
        self.desc_edit = QtWidgets.QLineEdit()
        self.template_edit = QtWidgets.QPlainTextEdit()
        form.addRow("Name:", self.name_edit)
        form.addRow("Description:", self.desc_edit)
        form.addRow("Template (use {placeholders}):", self.template_edit)
        layout.addLayout(form)
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        save_btn = btns.button(QtWidgets.QDialogButtonBox.Save)
        cancel_btn = btns.button(QtWidgets.QDialogButtonBox.Cancel)
        if save_btn is not None:
            save_btn.clicked.connect(self.on_save)
        else:
            btns.accepted.connect(self.on_save)
        if cancel_btn is not None:
            cancel_btn.clicked.connect(self.reject)
        else:
            btns.rejected.connect(self.reject)
        layout.addWidget(btns)
        if procedure:
            self.name_edit.setText(procedure.get("name", ""))
            self.desc_edit.setText(procedure.get("description", ""))
            self.template_edit.setPlainText(procedure.get("template", ""))
    def on_save(self):
        self.accept()
    def get_data(self):
        return {
            "name": self.name_edit.text().strip(),
            "description": self.desc_edit.text().strip(),
            "template": self.template_edit.toPlainText().strip()
        }

class ProcedureManagerDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, procedures=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Procedures")
        self.resize(800, 450)
        self.procedures = [dict(p) for p in (procedures or [])]
        layout = QtWidgets.QVBoxLayout(self)
        self.list_widget = QtWidgets.QListWidget()
        layout.addWidget(self.list_widget)
        btn_row = QtWidgets.QHBoxLayout()
        add_btn = QtWidgets.QPushButton("Add")
        edit_btn = QtWidgets.QPushButton("Edit")
        del_btn = QtWidgets.QPushButton("Delete")
        up_btn = QtWidgets.QPushButton("Move Up")
        down_btn = QtWidgets.QPushButton("Move Down")
        btn_row.addWidget(add_btn)
        btn_row.addWidget(edit_btn)
        btn_row.addWidget(del_btn)
        btn_row.addWidget(up_btn)
        btn_row.addWidget(down_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)
        add_btn.clicked.connect(self.add_procedure)
        edit_btn.clicked.connect(self.edit_procedure)
        del_btn.clicked.connect(self.delete_procedure)
        up_btn.clicked.connect(self.move_up)
        down_btn.clicked.connect(self.move_down)
        bottom = QtWidgets.QHBoxLayout()
        save_btn = QtWidgets.QPushButton("Save && Close")
        cancel_btn = QtWidgets.QPushButton("Cancel")
        bottom.addStretch()
        bottom.addWidget(save_btn)
        bottom.addWidget(cancel_btn)
        layout.addLayout(bottom)
        save_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        self.refresh_list()
    def refresh_list(self):
        self.list_widget.clear()
        for p in self.procedures:
            name = p.get("name", "<no name>")
            desc = p.get("description", "")
            item = QtWidgets.QListWidgetItem(f"{name} — {desc}")
            self.list_widget.addItem(item)
    def add_procedure(self):
        dlg = ProcedureEditorDialog(self, procedure=None)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            data = dlg.get_data()
            pid = f"proc{int(time.time() * 1000)}"
            entry = {"id": pid, "name": data.get("name", ""), "description": data.get("description", ""), "template": data.get("template", "")}
            self.procedures.append(entry)
            self.refresh_list()
    def edit_procedure(self):
        item = self.list_widget.currentItem()
        if not item:
            return
        idx = self.list_widget.currentRow()
        proc = self.procedures[idx]
        dlg = ProcedureEditorDialog(self, procedure=proc)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            data = dlg.get_data()
            proc["name"] = data.get("name", proc.get("name", ""))
            proc["description"] = data.get("description", proc.get("description", ""))
            proc["template"] = data.get("template", proc.get("template", ""))
            self.procedures[idx] = proc
            self.refresh_list()
    def delete_procedure(self):
        item = self.list_widget.currentItem()
        if not item:
            return
        idx = self.list_widget.currentRow()
        proc = self.procedures[idx]
        confirm = QtWidgets.QMessageBox.question(self, "Delete", f"Delete procedure '{proc.get('name', '')}'?")
        if confirm == QtWidgets.QMessageBox.StandardButton.Yes:
            del self.procedures[idx]
            self.refresh_list()
    def move_up(self):
        idx = self.list_widget.currentRow()
        if idx <= 0:
            return
        self.procedures[idx - 1], self.procedures[idx] = self.procedures[idx], self.procedures[idx - 1]
        self.refresh_list()
        self.list_widget.setCurrentRow(idx - 1)
    def move_down(self):
        idx = self.list_widget.currentRow()
        if idx < 0 or idx >= len(self.procedures) - 1:
            return
        self.procedures[idx + 1], self.procedures[idx] = self.procedures[idx], self.procedures[idx + 1]
        self.refresh_list()
        self.list_widget.setCurrentRow(idx + 1)

class TrivialNamesManagerDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, mapping=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Trivial Names")
        self.resize(600, 400)
        self.mapping = dict(mapping or {})
        layout = QtWidgets.QVBoxLayout(self)
        self.list_widget = QtWidgets.QListWidget()
        layout.addWidget(self.list_widget)
        btn_layout = QtWidgets.QHBoxLayout()
        add_btn = QtWidgets.QPushButton("Add")
        edit_btn = QtWidgets.QPushButton("Edit")
        del_btn = QtWidgets.QPushButton("Delete")
        import_btn = QtWidgets.QPushButton("Import JSON...")
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(del_btn)
        btn_layout.addWidget(import_btn)
        layout.addLayout(btn_layout)
        add_btn.clicked.connect(self.add_entry)
        edit_btn.clicked.connect(self.edit_entry)
        del_btn.clicked.connect(self.delete_entry)
        import_btn.clicked.connect(self.import_json)
        close_btn = QtWidgets.QPushButton("Save && Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
        self.refresh_list()
    def refresh_list(self):
        self.list_widget.clear()
        for k, v in sorted(self.mapping.items(), key=lambda x: x[0]):
            disp_key = subscript_digits(k)
            self.list_widget.addItem(f"{disp_key} → {v}")
    def add_entry(self):
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Add trivial name")
        form = QtWidgets.QFormLayout(dlg)
        formula_edit = QtWidgets.QLineEdit()
        name_edit = QtWidgets.QLineEdit()
        form.addRow("Formula (ASCII, e.g. C18H20):", formula_edit)
        form.addRow("Trivial name:", name_edit)
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        form.addWidget(btns)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            f = formula_edit.text().strip()
            n = name_edit.text().strip()
            if f and n:
                f_norm = remove_subscript_unicode(strip_markup_to_plain_ascii(f)).strip()
                self.mapping[f_norm] = n
                self.refresh_list()
    def edit_entry(self):
        item = self.list_widget.currentItem()
        if not item:
            return
        idx = self.list_widget.currentRow()
        key = sorted(self.mapping.keys())[idx]
        val = self.mapping.get(key, "")
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Edit trivial name")
        form = QtWidgets.QFormLayout(dlg)
        formula_edit = QtWidgets.QLineEdit(key)
        name_edit = QtWidgets.QLineEdit(val)
        form.addRow("Formula (ASCII):", formula_edit)
        form.addRow("Trivial name:", name_edit)
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        form.addWidget(btns)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            f = formula_edit.text().strip()
            n = name_edit.text().strip()
            if f:
                f_norm = remove_subscript_unicode(strip_markup_to_plain_ascii(f)).strip()
                if f_norm != key and key in self.mapping:
                    del self.mapping[key]
                self.mapping[f_norm] = n
                self.refresh_list()
    def delete_entry(self):
        item = self.list_widget.currentItem()
        if not item:
            return
        idx = self.list_widget.currentRow()
        key = sorted(self.mapping.keys())[idx]
        confirm = QtWidgets.QMessageBox.question(self, "Delete", f"Delete entry for '{key}'?")
        if confirm == QtWidgets.QMessageBox.StandardButton.Yes:
            del self.mapping[key]
            self.refresh_list()
    def import_json(self):
        fpath, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Import trivial names JSON", str(Path.cwd()), "JSON files (*.json);;All files (*)")
        if not fpath:
            return
        try:
            with open(fpath, "r", encoding="utf-8") as fh:
                data = json.load(fh)
                if isinstance(data, dict):
                    for k, v in data.items():
                        kk = remove_subscript_unicode(strip_markup_to_plain_ascii(k)).strip()
                        if kk:
                            self.mapping[kk] = str(v)
                    self.refresh_list()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to import: {e}")

class EditOutputDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, initial_markup=""):
        super().__init__(parent)
        self.setWindowTitle("Edit Output (temporary)")
        self.resize(700, 400)
        layout = QtWidgets.QVBoxLayout(self)
        self.editor = QtWidgets.QPlainTextEdit()
        self.editor.setPlainText(initial_markup or "")
        # Increase editor font size by 3 points relative to preview font if available
        try:
            font = self.editor.font()
            base_pt = None
            if parent is not None and hasattr(parent, "preview_font_pt") and isinstance(parent.preview_font_pt, (int, float)):
                base_pt = int(parent.preview_font_pt)
            else:
                # fallback to application's default font point size or editor's current size
                base_pt = int(font.pointSize()) if font.pointSize() > 0 else int(QtWidgets.QApplication.font().pointSize() or 12)
            font.setPointSize(max(1, base_pt + 3))
            self.editor.setFont(font)
        except Exception:
            # non-fatal: ignore font adjustments on failure
            pass
        layout.addWidget(self.editor)
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)
    def get_text(self):
        return self.editor.toPlainText()

class AssignedMappingDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, placeholders=None, mapping=None):
        super().__init__(parent)
        self.setWindowTitle("Assigned Mapping")
        self.resize(500, 350)
        self.placeholders = list(placeholders or [])
        self.mapping = dict(mapping or {})
        layout = QtWidgets.QVBoxLayout(self)
        self.list_widget = QtWidgets.QListWidget()
        layout.addWidget(self.list_widget)
        btn_row = QtWidgets.QHBoxLayout()
        unassign_btn = QtWidgets.QPushButton("Unassign Selected")
        copy_btn = QtWidgets.QPushButton("Copy to Clipboard")
        btn_row.addWidget(unassign_btn)
        btn_row.addWidget(copy_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)
        close_btn = QtWidgets.QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
        unassign_btn.clicked.connect(self.unassign_selected)
        copy_btn.clicked.connect(self.copy_to_clipboard)
        self.refresh_list()
    def refresh_list(self):
        self.list_widget.clear()
        for ph in self.placeholders:
            val = self.mapping.get(ph, "")
            li = QtWidgets.QListWidgetItem(f"{ph} → {val or '<unassigned>'}")
            self.list_widget.addItem(li)
    def unassign_selected(self):
        item = self.list_widget.currentItem()
        if not item:
            return
        idx = self.list_widget.currentRow()
        ph = self.placeholders[idx]
        if ph in self.mapping:
            del self.mapping[ph]
        self.refresh_list()
    def copy_to_clipboard(self):
        lines = []
        for ph in self.placeholders:
            val = self.mapping.get(ph, "")
            lines.append(f"{ph} -> {val or ''}")
        cb = QtWidgets.QApplication.clipboard()
        cb.setText("\n".join(lines))

class LearningManagementDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, learning=None, procedures=None):
        super().__init__(parent)
        self.setWindowTitle("Learning Management")
        self.resize(760, 420)
        self.learning = learning if isinstance(learning, dict) else {}
        self.procedure_names = {
            str(procedure.get("id", "")): procedure.get("name", "<unnamed procedure>")
            for procedure in (procedures or [])
        }
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(QtWidgets.QLabel("Learned reaction and chemical assignments"))
        self.list_widget = QtWidgets.QListWidget()
        layout.addWidget(self.list_widget, 1)
        button_row = QtWidgets.QHBoxLayout()
        delete_btn = QtWidgets.QPushButton("Delete Selected")
        reset_btn = QtWidgets.QPushButton("Reset All")
        button_row.addWidget(delete_btn)
        button_row.addWidget(reset_btn)
        button_row.addStretch()
        layout.addLayout(button_row)
        close_btn = QtWidgets.QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
        delete_btn.clicked.connect(self.delete_selected)
        reset_btn.clicked.connect(self.reset_all)
        self.refresh_list()

    def refresh_list(self):
        self.list_widget.clear()
        for signature, record in sorted(self.learning.items()):
            if not isinstance(record, dict):
                continue
            procedure_ids = set(record.get("__mapping_by_procedure__", {}))
            procedure_ids.update(key for key in record if not key.startswith("__"))
            for procedure_id in sorted(procedure_ids):
                mapping = record.get("__mapping_by_procedure__", {}).get(procedure_id, {})
                reaction_score = record.get(procedure_id, 0)
                procedure_name = self.procedure_names.get(procedure_id, procedure_id)
                roles = ", ".join(f"{key}={value}" for key, value in sorted(mapping.items()))
                label = f"{procedure_name} | {signature} | score {reaction_score}"
                if roles:
                    label += f" | {roles}"
                item = QtWidgets.QListWidgetItem(label)
                item.setData(QtCore.Qt.UserRole, (signature, procedure_id))
                self.list_widget.addItem(item)

    def delete_selected(self):
        item = self.list_widget.currentItem()
        if not item:
            return
        signature, procedure_id = item.data(QtCore.Qt.UserRole)
        record = self.learning.get(signature, {})
        if isinstance(record, dict):
            record.pop(procedure_id, None)
            mappings = record.get("__mapping_by_procedure__", {})
            if isinstance(mappings, dict):
                mappings.pop(procedure_id, None)
                if not mappings:
                    record.pop("__mapping_by_procedure__", None)
            if not any(not key.startswith("__") for key in record):
                self.learning.pop(signature, None)
        self.refresh_list()

    def reset_all(self):
        if not self.learning:
            return
        confirm = QtWidgets.QMessageBox.question(
            self, "Reset learning", "Delete all learned assignments and reaction history?"
        )
        if confirm == QtWidgets.QMessageBox.StandardButton.Yes:
            self.learning.clear()
            self.refresh_list()

class DropArea(QtWidgets.QWidget):
    filesDropped = QtCore.Signal(list)
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setAutoFillBackground(True)
        pal = self.palette()
        self.setPalette(pal)
    def dragEnterEvent(self, event):
        md = event.mimeData()
        if md.hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()
    def dragMoveEvent(self, event):
        md = event.mimeData()
        if md.hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()
    def dropEvent(self, event):
        md = event.mimeData()
        if not md.hasUrls():
            event.ignore()
            return
        paths = []
        for url in md.urls():
            local = url.toLocalFile()
            if local:
                paths.append(local)
        if paths:
            self.filesDropped.emit(paths)
            event.acceptProposedAction()
        else:
            event.ignore()

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Experimental Section Writer")
        self.resize(1100, 760)
        self.procedures = load_procedures()
        self.solvents = load_solvents()
        self.trivial_names = load_trivial_names()
        self.reaction_learning = load_reaction_learning()
        self.chemicals = []
        self.mapping = {}
        self.automatic_mapping = set()
        self.last_rendered_markup = ""
        self.current_template = ""
        self._preview_internal_update = False
        self._saved_preview_cursor = None
        self.preview_font_pt = None
        # history of (mapping, last_rendered_markup)
        self.state_history = []
        self.history_limit = 50
        self.setup_ui()

    def push_state(self):
        # store a deep copy of mapping and the last rendered markup
        snap = (deepcopy(self.mapping), str(self.last_rendered_markup))
        self.state_history.append(snap)
        if len(self.state_history) > self.history_limit:
            self.state_history.pop(0)

    def undo_last_action(self):
        if not self.state_history:
            QtWidgets.QMessageBox.information(self, "Undo", "Nothing to undo.")
            return
        last_map, last_markup = self.state_history.pop()
        self.mapping = deepcopy(last_map)
        self.last_rendered_markup = str(last_markup)
        # update preview based on restored markup
        self.update_preview_from_markup(self.last_rendered_markup)
        self.refresh_mapping_list()
        self.statusBar().showMessage("Undo applied", 3000)

    def setup_ui(self):
        men = self.menuBar()
        file_menu = men.addMenu("File")
        load_act = QtGui.QAction("Load CDMXL...", self)
        load_act.triggered.connect(self.load_cdmxl)
        file_menu.addAction(load_act)
        procedures_menu = men.addMenu("Procedures")
        manage_act = QtGui.QAction("Manage Procedures...", self)
        manage_act.triggered.connect(self.open_procedure_manager)
        procedures_menu.addAction(manage_act)
        learning_menu = men.addMenu("Learning")
        manage_learning_act = QtGui.QAction("Manage Learned Assignments...", self)
        manage_learning_act.triggered.connect(self.open_learning_manager)
        learning_menu.addAction(manage_learning_act)
        solvents_act = QtGui.QAction("Solvents", self)
        solvents_act.triggered.connect(self.open_solvent_manager)
        men.addAction(solvents_act)
        names_menu = men.addMenu("Names")
        manage_triv_act = QtGui.QAction("Manage Trivial Names...", self)
        manage_triv_act.triggered.connect(self.open_trivial_names_manager)
        names_menu.addAction(manage_triv_act)
        view_menu = men.addMenu("View")
        assigned_mapping_act = QtGui.QAction("Assigned Mapping...", self)
        assigned_mapping_act.triggered.connect(self.open_assigned_mapping_dialog)
        view_menu.addAction(assigned_mapping_act)

        central = DropArea(self)
        central.filesDropped.connect(self.load_cdxml_paths)
        self.setCentralWidget(central)
        layout = QtWidgets.QHBoxLayout(central)
        layout.setContentsMargins(8, 8, 8, 8)

        left_frame = QtWidgets.QFrame()
        left_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        left_frame.setMinimumWidth(300)
        left_layout = QtWidgets.QVBoxLayout(left_frame)
        left_layout.setSpacing(10)

        left_layout.addWidget(QtWidgets.QLabel("Chemicals (from CDX/CDXML)"))
        self.chem_list = DraggableListWidget()
        left_layout.addWidget(self.chem_list, 1)

        left_layout.addWidget(QtWidgets.QLabel("Solvents"))
        self.solv_list = DraggableListWidget()
        left_layout.addWidget(self.solv_list, 2)

        # Bottom-left Load button removed per request; keep only Manage button
        left_btn_row = QtWidgets.QHBoxLayout()
        manage_solvents_btn = QtWidgets.QPushButton("Manage")
        manage_solvents_btn.setProperty("secondary", True)
        manage_solvents_btn.clicked.connect(self.open_solvent_manager)
        left_btn_row.addWidget(manage_solvents_btn)
        left_layout.addLayout(left_btn_row)

        right_frame = QtWidgets.QFrame()
        right_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        right_layout = QtWidgets.QVBoxLayout(right_frame)
        right_layout.setSpacing(8)

        proc_row = QtWidgets.QHBoxLayout()
        proc_row.addWidget(QtWidgets.QLabel("Procedure:"))
        self.proc_combo = QtWidgets.QComboBox()
        self.proc_combo.setMaxVisibleItems(12)
        # Verbindung hinzufügen: bei Auswahländerung sofort on_proc_changed aufrufen
        self.proc_combo.currentIndexChanged.connect(self.on_proc_changed)
        proc_row.addWidget(self.proc_combo, 1)

        load_btn2 = QtWidgets.QPushButton("Load")
        load_btn2.setToolTip("Load CDMXL files or folders")
        load_btn2.clicked.connect(self.load_cdmxl)
        proc_row.addWidget(load_btn2)

        right_layout.addLayout(proc_row)

        preview_lbl = QtWidgets.QLabel("Preview")
        right_layout.addWidget(preview_lbl)

        format_row = QtWidgets.QHBoxLayout()
        format_row.setSpacing(4)
        self.bold_btn = self._create_format_button("B", "Bold", self.toggle_bold)
        self.italic_btn = self._create_format_button("I", "Italic", self.toggle_italic)
        self.superscript_btn = self._create_format_button("x²", "Superscript", self.toggle_superscript)
        self.subscript_btn = self._create_format_button("x₂", "Subscript", self.toggle_subscript)
        format_row.addStretch(1)
        for button in (self.bold_btn, self.italic_btn, self.superscript_btn, self.subscript_btn):
            format_row.addWidget(button)
        right_layout.addLayout(format_row)

        self.preview = PreviewText()
        font = self.preview.font()
        base = font.pointSize()
        if base <= 0:
            base = QtWidgets.QApplication.font().pointSize()
            if base <= 0:
                base = 12
        preview_size = max(base + 5, 14)
        font.setPointSize(preview_size)
        self.preview.setFont(font)
        self.preview_font_pt = preview_size
        # ensure objectNames for styled buttons (set later when creating buttons)
        self.preview.dropped.connect(self.on_preview_dropped)
        self.preview.drag_hovered.connect(self.on_preview_drag_hovered)
        self.preview.drag_left.connect(self.clear_preview_drag_hover)
        self.preview.textChanged.connect(self.on_preview_text_edited)
        self.preview.cursorPositionChanged.connect(self.update_format_buttons)
        self.preview.selectionChanged.connect(self.update_format_buttons)
        right_layout.addWidget(self.preview, 1)

        bottom_h = QtWidgets.QHBoxLayout()
        preview_btn = QtWidgets.QPushButton("Render Preview")
        preview_btn.clicked.connect(self.render_preview)
        bottom_h.addWidget(preview_btn)

        # Undo button (objectName used for styling)
        undo_btn = QtWidgets.QPushButton("Undo")
        undo_btn.setObjectName("undoBtn")
        undo_btn.clicked.connect(self.undo_last_action)
        bottom_h.addWidget(undo_btn)

        copy_preview_btn = QtWidgets.QPushButton("Copy Preview")
        copy_preview_btn.setObjectName("copyPreviewBtn")
        copy_preview_btn.clicked.connect(self.copy_preview_to_clipboard)
        bottom_h.addWidget(copy_preview_btn)

        export_word_btn = QtWidgets.QPushButton("Export to Word")
        export_word_btn.clicked.connect(self.export_word)
        bottom_h.addWidget(export_word_btn)

        edit_output_btn = QtWidgets.QPushButton("Edit Output")
        # set objectName for centralized styling
        edit_output_btn.setObjectName("editOutputBtn")
        edit_output_btn.setToolTip("Edit the generated output text (temporary).")
        edit_output_btn.clicked.connect(self.open_edit_output_dialog)
        bottom_h.addWidget(edit_output_btn)

        convert_btn = QtWidgets.QPushButton("Convert Names")
        convert_btn.setToolTip("Convert chemical formulas in the preview using trivial_names.json only (no network).")
        convert_btn.clicked.connect(self.convert_preview_formulas_to_names)
        bottom_h.addWidget(convert_btn)

        for b in (export_word_btn, edit_output_btn):
            b.setProperty("secondary", False)

        right_layout.addLayout(bottom_h)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        splitter.addWidget(left_frame)
        splitter.addWidget(right_frame)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        layout.addWidget(splitter, 1)

        status = QtWidgets.QLabel("")
        self.statusBar().addWidget(status)

        self.reload_procedures()
        self.refresh_solvents_list()

    # ---------------------------
    # Loading CDXML files (unchanged)
    # ---------------------------

    def load_cdmxl(self):
        fpath, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Open ChemDraw file (or cancel to choose folder)", str(Path.cwd()),
            "ChemDraw Files (*.cdx *.cdxml *.xml);;All files (*)"
        )
        if fpath:
            self.load_cdxml_paths([fpath])
            return
        return

    def load_cdxml_paths(self, paths):
    # Wenn neue Dateien/Folders geladen werden: vorherige Zuordnungen rücksetzen
    # und Preview auf ungefüllten Template-Zustand bringen.
        try:
            # Falls es aktuelle Zuordnungen oder gerenderten Text gibt, in die History pushen
            # damit das Löschen rückgängig gemacht werden kann.
            try:
                if getattr(self, "mapping", None) or getattr(self, "last_rendered_markup", None):
                    self.push_state()
            except Exception:
                # im Fehlerfall trotzdem fortfahren
                pass

            # Entferne alle aktuellen Assignments (keine Undo-Pushes - bereits oben gepusht)
            self.mapping = {}
            self.automatic_mapping.clear()
            self.last_rendered_markup = ""
            # Clear preview widget content
            if hasattr(self, "preview") and self.preview is not None:
                try:
                    self.preview.clear()
                except Exception:
                    pass
            # Clear chemicals list UI while we load new ones
            if hasattr(self, "chem_list") and self.chem_list is not None:
                try:
                    self.chem_list.clear()
                except Exception:
                    pass
            # Clear placeholders list UI (will be repopulated if procedure selected)
            if hasattr(self, "placeholders_list") and self.placeholders_list is not None:
                try:
                    self.placeholders_list.clear()
                except Exception:
                    pass
            # Update status of mapping/UI
            try:
                self.refresh_mapping_list()
            except Exception:
                pass
        except Exception:
            # be robust: don't block loading if clearing fails
            pass

        to_parse = []
        for p in paths:
            if os.path.isfile(p):
                to_parse.append(p)
            elif os.path.isdir(p):
                found = find_cdxml_files(p)
                if found:
                    to_parse.extend(found)
        to_parse = list(dict.fromkeys(to_parse))
        if not to_parse:
            QtWidgets.QMessageBox.information(self, "No files", "No CDX/CDXML/XML files found in the dropped items.")
            return
        components_all = []
        for file in to_parse:
            try:
                comps = parse_stoichiometry_components(file)
            except Exception:
                comps = []
            if comps:
                components_all.extend(comps)
        candidates = []
        for idx, comp in enumerate(components_all):
            is_last = (idx == len(components_all) - 1)
            formatted = format_component_display_by_index(comp, is_last=is_last)
            if not formatted["formula"]:
                continue
            candidates.append({
                "formula": formatted["formula"],
                "display": formatted["display"]
            })
        seen = set()
        normalized = []
        for c in candidates:
            f = c["formula"].strip()
            if not f:
                continue
            key = f.lower()
            if key in seen:
                continue
            seen.add(key)
            normalized.append(c)
        if not normalized:
            QtWidgets.QMessageBox.information(self, "No chemicals", "No candidate chemicals found in the selected file(s).")
            # preview is already cleared at the top; ensure UI mapping state consistent
            try:
                self.refresh_mapping_list()
            except Exception:
                pass
            return
        self.chemicals = normalized
        # populate chemicals list UI
        self.chem_list.clear()
        for c in self.chemicals:
            self.chem_list.addItem(c["display"])
        self.statusBar().showMessage(f"Loaded {len(self.chemicals)} candidate chemicals", 5000)

        suggested_index = suggest_reaction_index(
            self.chemicals,
            self.procedures,
            self.reaction_learning,
        )
        if suggested_index is not None and suggested_index != self.proc_combo.currentIndex():
            self.proc_combo.setCurrentIndex(suggested_index)
            suggested_name = self.proc_combo.itemText(suggested_index)
            self.statusBar().showMessage(f"Suggested reaction: {suggested_name}", 6000)

        # Ensure assignments remain cleared and preview shows the template with placeholders unfilled.
        # (If a procedure/template is selected, render it with placeholders unassigned)
        try:
            self.mapping = {}
            self.apply_automatic_mapping()
            # render the current template (will show placeholders as {name} if unassigned)
            self.render_preview()
            self.refresh_mapping_list()
        except Exception:
            # non-fatal: ignore rendering errors here
            pass

    def refresh_solvents_list(self):
        self.solv_list.clear()
        for s in self.solvents:
            self.solv_list.addItem(subscript_digits(s))

    def open_solvent_manager(self):
        dlg = SolventManagerDialog(self, solvents=self.solvents)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            self.solvents = dlg.solvents
            save_solvents(self.solvents)
            self.refresh_solvents_list()

    def reload_procedures(self):
        self.proc_combo.clear()
        for p in self.procedures:
            name = p.get("name") or "<no name>"
            self.proc_combo.addItem(name, p)
        if hasattr(self, "placeholders_list") and hasattr(self, "proc_combo") and self.procedures:
            self.proc_combo.setCurrentIndex(0)
            # On startup select first and render preview (on_proc_changed will call render)
            self.on_proc_changed(0)

    def open_procedure_manager(self):
        dlg = ProcedureManagerDialog(self, procedures=self.procedures)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            self.procedures = dlg.procedures
            save_procedures(self.procedures)
            self.reload_procedures()

    def open_learning_manager(self):
        dlg = LearningManagementDialog(
            self,
            learning=self.reaction_learning,
            procedures=self.procedures,
        )
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            save_reaction_learning(self.reaction_learning)
            self.statusBar().showMessage("Learning data updated", 3000)

    def on_proc_changed(self, idx):
        p = self.proc_combo.currentData()
        if not p:
            if hasattr(self, "placeholders_list"):
                self.placeholders_list.clear()
            return
        self.current_template = p.get("template", "")
        phs = extract_placeholders(self.current_template)
        self.placeholders = phs
        self.mapping = {}
        self.automatic_mapping.clear()
        self.apply_automatic_mapping()
        self.refresh_placeholders_list()
        # render preview immediately
        self.render_preview()

    def apply_automatic_mapping(self):
        procedure = self.proc_combo.currentData()
        automatic = suggest_chemical_mapping(
            getattr(self, "chemicals", []),
            getattr(self, "placeholders", []),
            procedure,
            self.reaction_learning,
        )
        self.mapping.update(automatic)
        self.automatic_mapping.update(automatic)
        self.refresh_placeholders_list()

    def refresh_placeholders_list(self):
        if not hasattr(self, "placeholders_list"):
            return
        self.placeholders_list.clear()
        for placeholder in getattr(self, "placeholders", []) or []:
            item = QtWidgets.QListWidgetItem(f"{{{placeholder}}}")
            item.setData(QtCore.Qt.UserRole, placeholder)
            self.placeholders_list.addItem(item)

    def learn_chemical_mapping(self, signature, procedure_id):
        if not signature or not procedure_id:
            return False
        assigned = {
            placeholder: value
            for placeholder, value in self.mapping.items()
            if placeholder in getattr(self, "placeholders", [])
            and value
            and "solvent" not in placeholder.lower()
        }
        if not assigned:
            return False
        learned_for_signature = self.reaction_learning.setdefault(signature, {})
        mappings_by_procedure = learned_for_signature.setdefault("__mapping_by_procedure__", {})
        learned_mapping = mappings_by_procedure.setdefault(str(procedure_id), {})
        changed = False
        for placeholder, value in assigned.items():
            if learned_mapping.get(placeholder) != value:
                learned_mapping[placeholder] = value
                changed = True
        if changed:
            save_reaction_learning(self.reaction_learning)
        return changed

    def on_placeholder_clicked(self, item):
        placeholder = item.data(QtCore.Qt.UserRole) or item.text().strip("{}")

        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle(f"Assign chemical to {{{placeholder}}}")
        dlg.resize(400, 300)
        layout = QtWidgets.QVBoxLayout(dlg)

        listw = QtWidgets.QListWidget()
        layout.addWidget(listw)

        # UNASSIGN option
        listw.addItem("<unassigned>")

        # Fill list depending on placeholder type
        if "solvent" in placeholder.lower():
            # solvents first
            for s in self.solvents:
                listw.addItem(subscript_digits(s))

            # separator
            sep = QtWidgets.QListWidgetItem("── Chemicals ──")
            sep.setFlags(QtCore.Qt.NoItemFlags)
            listw.addItem(sep)

            for c in self.chemicals:
                listw.addItem(c.get("display"))
        else:
            # chemicals first
            for c in self.chemicals:
                listw.addItem(c.get("display"))

            sep = QtWidgets.QListWidgetItem("── Solvents ──")
            sep.setFlags(QtCore.Qt.NoItemFlags)
            listw.addItem(sep)

            for s in self.solvents:
                listw.addItem(subscript_digits(s))

        # Click behavior — immediate selection + close dialog
        def on_item_clicked(clicked_item):
            text = clicked_item.text()

            # push current state for undo
            self.push_state()

            if text == "<unassigned>":
                if placeholder in self.mapping:
                    del self.mapping[placeholder]
            elif "──" not in text:  # ignore separators
                self.mapping[placeholder] = text
                self.automatic_mapping.discard(placeholder)

            self.render_preview()
            self.refresh_placeholders_list()
            dlg.accept()  # close instantly

        listw.itemClicked.connect(on_item_clicked)

        dlg.exec()

    def refresh_mapping_list(self):
        assigned = sum(1 for v in self.mapping.values() if v)
        self.statusBar().showMessage(f"{assigned}/{len(getattr(self, 'placeholders', []) or [])} placeholders assigned", 4000)

    def render_preview(self):
        if not hasattr(self, "current_template"):
            return
        rendered = self._render_markup(self.mapping)
        self.last_rendered_markup = rendered
        self.update_preview_from_markup(rendered, highlight_mapping=self.mapping)
        self.refresh_mapping_list()

    def _render_markup(self, mapping):
        text = self.current_template
        def repl(m):
            key = m.group(1)
            return mapping.get(key, f"{{{key}}}")
        return re.sub(r"\{([a-zA-Z0-9_]+)\}", repl, text)

    def update_preview_from_markup(self, markup, highlight_mapping=None):
        if hasattr(self, "preview"):
            self._preview_internal_update = True
            try:
                self.preview.setHtml(
                    self._build_preview_html(
                        markup,
                        word_compatible=False,
                        highlight_mapping=highlight_mapping,
                    )
                )
            finally:
                self._preview_internal_update = False

    def _preview_document_to_markup(self):
        lines = []
        block = self.preview.document().begin()
        while block.isValid():
            parts = []
            fragment = block.begin()
            while fragment != block.end():
                text = fragment.fragment().text()
                char_format = fragment.fragment().charFormat()
                if char_format.verticalAlignment() == QtGui.QTextCharFormat.AlignSubScript:
                    text = f"<sub>{text}</sub>"
                if char_format.fontItalic():
                    text = f"<i>{text}</i>"
                if char_format.fontWeight() >= QtGui.QFont.Bold:
                    text = f"<b>{text}</b>"
                parts.append(text)
                fragment += 1
            lines.append("".join(parts))
            block = block.next()
        return "\n".join(lines)

    def on_preview_text_edited(self):
        if self._preview_internal_update:
            return
        self.last_rendered_markup = self._preview_document_to_markup()

    def _create_format_button(self, label, tooltip, slot):
        button = QtWidgets.QToolButton()
        button.setObjectName("formatToolButton")
        button.setText(label)
        button.setToolTip(tooltip)
        button.setCheckable(True)
        if label == "I":
            font = button.font()
            font.setItalic(True)
            button.setFont(font)
        button.pressed.connect(self.remember_preview_cursor)
        button.clicked.connect(slot)
        return button

    def remember_preview_cursor(self):
        if hasattr(self, "preview"):
            self._saved_preview_cursor = QtGui.QTextCursor(self.preview.textCursor())

    def _merge_preview_format(self, char_format):
        cursor = self._saved_preview_cursor or self.preview.textCursor()
        if cursor.hasSelection():
            cursor.mergeCharFormat(char_format)
            self.preview.setTextCursor(cursor)
        else:
            self.preview.mergeCurrentCharFormat(char_format)
        self._saved_preview_cursor = None
        self.update_format_buttons()

    def _toggle_preview_format(self, kind, checked):
        char_format = QtGui.QTextCharFormat()
        if kind == "bold":
            char_format.setFontWeight(QtGui.QFont.Bold if checked else QtGui.QFont.Normal)
        elif kind == "italic":
            char_format.setFontItalic(checked)
        else:
            char_format.setVerticalAlignment(
                QtGui.QTextCharFormat.AlignSuperScript
                if kind == "superscript" and checked
                else QtGui.QTextCharFormat.AlignSubScript
                if kind == "subscript" and checked
                else QtGui.QTextCharFormat.AlignNormal
            )
        if checked and kind == "superscript":
            self.subscript_btn.setChecked(False)
        elif checked and kind == "subscript":
            self.superscript_btn.setChecked(False)
        self._merge_preview_format(char_format)

    def toggle_bold(self, checked=False):
        self._toggle_preview_format("bold", checked)

    def toggle_italic(self, checked=False):
        self._toggle_preview_format("italic", checked)

    def toggle_superscript(self, checked=False):
        self._toggle_preview_format("superscript", checked)

    def toggle_subscript(self, checked=False):
        self._toggle_preview_format("subscript", checked)

    def update_format_buttons(self):
        if not hasattr(self, "preview"):
            return
        char_format = self.preview.textCursor().charFormat()
        self.bold_btn.setChecked(char_format.fontWeight() >= QtGui.QFont.Bold)
        self.italic_btn.setChecked(char_format.fontItalic())
        alignment = char_format.verticalAlignment()
        self.superscript_btn.setChecked(alignment == QtGui.QTextCharFormat.AlignSuperScript)
        self.subscript_btn.setChecked(alignment == QtGui.QTextCharFormat.AlignSubScript)

    def open_edit_output_dialog(self):
        if not getattr(self, "last_rendered_markup", "").strip():
            self.render_preview()
        initial = getattr(self, "last_rendered_markup", "") or ""
        dlg = EditOutputDialog(self, initial_markup=initial)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            # push previous state for undo
            self.push_state()
            new_text = dlg.get_text() or ""
            self.last_rendered_markup = new_text
            self.update_preview_from_markup(new_text)

    def _convert_unicode_subs_to_html_subs(self, s: str) -> str:
        if not s:
            return s
        chars = "".join(re.escape(ch) for ch in UNICODE_SUBSCRIPT_DIGITS)
        pattern = re.compile(f"[{chars}]+")
        def repl(m):
            uni = m.group(0)
            ascii_digits = uni.translate(SUBSCRIPT_REVERSE_TRANSLATION)
            return f"<sub>{ascii_digits}</sub>"
        return pattern.sub(repl, s)

    def _build_preview_html(self, markup: str, word_compatible: bool = False, highlight_mapping=None) -> str:
        markup_for_html = self._convert_unicode_subs_to_html_subs(markup or "")
        paragraphs = []
        font_pt = self.preview_font_pt or 14
        font_style = "" if word_compatible else "font-family:Calibri, Arial, sans-serif; "
        for line in markup_for_html.splitlines() or [""]:
            frag = render_markup_to_html(line)
            if highlight_mapping is not None and not word_compatible:
                for placeholder, value in sorted(
                    highlight_mapping.items(), key=lambda item: len(str(item[1])), reverse=True
                ):
                    if not value:
                        continue
                    value_html = render_markup_to_html(
                        self._convert_unicode_subs_to_html_subs(str(value))
                    )
                    if value_html not in frag:
                        continue
                    color = "#163d63" if placeholder in self.automatic_mapping else "#1d4930"
                    border = "#0a84ff" if placeholder in self.automatic_mapping else "#63d471"
                    highlight = (
                        f'<span style="background-color:{color}; border:1px solid {border}; '
                        f'border-radius:3px; padding:1px 3px;">{value_html}</span>'
                    )
                    frag = frag.replace(value_html, highlight)
                assigned_placeholders = {
                    placeholder for placeholder, value in highlight_mapping.items() if value
                }
                for placeholder in getattr(self, "placeholders", []) or []:
                    if placeholder in assigned_placeholders:
                        continue
                    token = html_std.escape(f"{{{placeholder}}}")
                    if token in frag:
                        frag = frag.replace(
                            token,
                            '<span style="background-color:#5a3a16; border:1px solid #f0a34a; '
                            f'border-radius:3px; padding:1px 3px; color:#ffd39a;">{token}</span>',
                        )
            paragraphs.append(
                f'<p style="{font_style}font-size:{font_pt}pt; '
                f'text-align:justify; line-height:1.5; margin:0 0 6px 0;">{frag}</p>'
            )
        body = "".join(paragraphs)
        if word_compatible:
            return (
                '<!DOCTYPE html><html><head><meta charset="utf-8"></head>'
                '<body><!--StartFragment-->' + body + '<!--EndFragment--></body></html>'
            )
        return "<html><body>" + body + "</body></html>"

    def _build_preview_rtf(self, markup: str) -> str:
        def escape_rtf(text: str) -> str:
            escaped = []
            for char in text:
                if char in "\\{}":
                    escaped.append("\\" + char)
                elif char == "\n":
                    escaped.append(r"\line ")
                elif ord(char) > 127:
                    escaped.append(f"\\u{ord(char)}?")
                else:
                    escaped.append(char)
            return "".join(escaped)

        parts = [r"{\rtf1\ansi\deff0\uc1{\fonttbl{\f0\fnil ;}} "]
        for line_index, line in enumerate((markup or "").splitlines() or [""]):
            last = 0
            for match in _MARKUP_RE.finditer(line):
                start, end = match.span()
                parts.append(escape_rtf(line[last:start]))
                if match.group(1):
                    inner = re.sub(r"(?i)^<sub>\s*|\s*</sub>$", "", match.group(1))
                    parts.append(r"{\sub " + escape_rtf(inner) + "}")
                elif match.group(2) or match.group(3):
                    inner = match.group(2) or match.group(3)
                    parts.append(r"{\sub " + escape_rtf(inner) + "}")
                elif match.group(4) or match.group(5):
                    inner = match.group(4) or match.group(5)
                    parts.append(r"{\i " + escape_rtf(inner) + "}")
                last = end
            parts.append(escape_rtf(line[last:]))
            if line_index < len((markup or "").splitlines()) - 1:
                parts.append(r"\line ")
        parts.append("}")
        return "".join(parts)

    def copy_preview_to_clipboard(self):
        self.convert_preview_formulas_to_names(show_success_popup=False)
        markup = getattr(self, "last_rendered_markup", "") or ""
        if not markup.strip():
            return
        html_full = self._build_preview_html(markup, word_compatible=True)
        plain_lines = []
        for line in markup.splitlines() or [""]:
            ascii_line = remove_subscript_unicode(line)
            ascii_line = strip_markup_to_plain_ascii(ascii_line)
            plain_lines.append(ascii_line)
        plain_ascii = "\n".join(plain_lines)
        mime = QtCore.QMimeData()
        mime.setHtml(html_full)
        mime.setData("text/rtf", QtCore.QByteArray(self._build_preview_rtf(markup).encode("ascii")))
        mime.setText(plain_ascii)
        cb = QtWidgets.QApplication.clipboard()
        cb.setMimeData(mime)

        procedure = self.proc_combo.currentData()
        procedure_id = procedure.get("id") if isinstance(procedure, dict) else None
        signature = chemical_signature(getattr(self, "chemicals", []))
        if procedure_id and signature:
            learned_for_signature = self.reaction_learning.setdefault(signature, {})
            learned_for_signature[procedure_id] = min(
                int(learned_for_signature.get(procedure_id, 0)) + 1,
                3,
            )
            save_reaction_learning(self.reaction_learning)
            mapping_changed = self.learn_chemical_mapping(signature, procedure_id)
            learned_message = " and chemical assignments" if mapping_changed else ""
            self.statusBar().showMessage(
                f"Reaction selection{learned_message} learned from copied preview", 4000
            )

    def export_word(self):
        markup = getattr(self, "last_rendered_markup", "") or ""
        if not markup.strip():
            QtWidgets.QMessageBox.warning(self, "Empty", "Nothing to export")
            return
        save_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Word Document", str(Path.cwd() / "experimental_section.docx"), "Word Document (*.docx)")
        if not save_path:
            return
        token_re = re.compile(
            r'(?P<htmlsub><sub>\s*([^<]+?)\s*</sub>)'
            r'|_(?P<brace_sub>\{([^}]+)\})'
            r'|_(?P<single_sub>[A-Za-z0-9])'
            r'|/(?P<brace_i>\{([^}]+)\})'
            r'|/(?P<single_i>[A-Za-z0-9])'
            r'|(?P<htmlbold><b>\s*([^<]+?)\s*</b>)',
            flags=re.IGNORECASE | re.DOTALL
        )
        try:
            doc = Document()
            for line in markup.splitlines() or [""]:
                p = doc.add_paragraph()
                pf = p.paragraph_format
                try:
                    pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                except Exception:
                    pass
                try:
                    pf.line_spacing = 1.5
                except Exception:
                    try:
                        pf.line_spacing = Pt(18)
                    except Exception:
                        pass
                last = 0
                for m in token_re.finditer(line):
                    if m.start() > last:
                        txt = line[last:m.start()]
                        p.add_run(txt)
                    if m.group('htmlbold'):
                        inner = re.sub(r'(?i)^<b>\s*', '', m.group('htmlbold'))
                        inner = re.sub(r'(?i)\s*</b>$', '', inner)
                        run = p.add_run(inner)
                        try:
                            run.bold = True
                        except Exception:
                            pass
                    elif m.group('htmlsub'):
                        inner = re.sub(r'(?i)^<sub>\s*', '', m.group('htmlsub'))
                        inner = re.sub(r'(?i)\s*</sub>$', '', inner)
                        run = p.add_run(inner)
                        try:
                            run.font.subscript = True
                        except Exception:
                            pass
                    elif m.group('brace_sub'):
                        inner = m.group(3)
                        run = p.add_run(inner)
                        try:
                            run.font.subscript = True
                        except Exception:
                            pass
                    elif m.group('single_sub'):
                        inner = m.group('single_sub')
                        run = p.add_run(inner)
                        try:
                            run.font.subscript = True
                        except Exception:
                            pass
                    elif m.group('brace_i'):
                        inner = m.group(6)
                        run = p.add_run(inner)
                        try:
                            run.font.italic = True
                        except Exception:
                            pass
                    elif m.group('single_i'):
                        inner = m.group('single_i')
                        run = p.add_run(inner)
                        try:
                            run.font.italic = True
                        except Exception:
                            pass
                    last = m.end()
                if last < len(line):
                    p.add_run(line[last:])
            doc.save(save_path)
            QtWidgets.QMessageBox.information(self, "Saved", f"Saved to {save_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to save Word: {e}")

    def open_trivial_names_manager(self):
        dlg = TrivialNamesManagerDialog(self, mapping=self.trivial_names)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            self.trivial_names = dlg.mapping
            save_trivial_names(self.trivial_names)
            self.statusBar().showMessage(f"Saved {len(self.trivial_names)} trivial name entries", 4000)

    def open_assigned_mapping_dialog(self):
        dlg = AssignedMappingDialog(self, placeholders=getattr(self, "placeholders", []), mapping=self.mapping)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            # push current state for undo, then copy returned mapping
            self.push_state()
            self.mapping = dlg.mapping
            self.automatic_mapping.difference_update(self.mapping)
            self.render_preview()
            self.refresh_placeholders_list()

    def convert_preview_formulas_to_names(self, show_success_popup=True):
        """
        Convert formulas in self.last_rendered_markup using only trivial_names mapping.
        No network lookups.
        """
        markup = getattr(self, "last_rendered_markup", "") or ""
        if not markup.strip():
            if show_success_popup:
                QtWidgets.QMessageBox.information(self, "Nothing", "No rendered text to convert.")
            return

        trivial_map = {k: v for k, v in self.trivial_names.items()}

        # find candidate tokens in markup that look like formulas:
        token_candidates = set(re.findall(r'[A-Za-z₀₁₂₃₄₅₆₇₈₉0-9\(\)]+', markup))
        candidates = []
        for tok in token_candidates:
            plain = remove_subscript_unicode(strip_markup_to_plain_ascii(tok)).strip()
            if not plain:
                continue
            if re.search(r'[A-Za-z]', plain) and re.search(r'\d', plain):
                candidates.append((tok, plain))

        if not candidates:
            if show_success_popup:
                QtWidgets.QMessageBox.information(self, "No formulas", "No formula-like tokens found to convert.")
            return

        # Sort to avoid partial replacements
        candidates.sort(key=lambda x: len(x[1]), reverse=True)

        new_markup = markup
        looked_up = {}

        # push previous state for undo
        self.push_state()

        for orig_tok, plain in candidates:
            plain_norm = plain.replace(" ", "")
            trivial_name = trivial_map.get(plain_norm)
            if trivial_name:
                new_markup = new_markup.replace(orig_tok, trivial_name)
                looked_up[plain_norm] = ("trivial", trivial_name)
            else:
                looked_up[plain_norm] = ("notfound", None)

        any_changed = new_markup != markup
        if any_changed:
            self.last_rendered_markup = new_markup
            self.update_preview_from_markup(new_markup)
            found_count = sum(1 for v in looked_up.values() if v[0] == "trivial")
            self.statusBar().showMessage(f"Converted {found_count} formula(s) using trivial names", 6000)
            if show_success_popup:
                QtWidgets.QMessageBox.information(self, "Converted", f"Converted {found_count} formula(s) in the preview.")
        else:
            if show_success_popup:
                QtWidgets.QMessageBox.information(self, "No changes", "No conversions were possible with trivial names.")

    # ---------------------------
    # Dropping on a rendered placeholder, including an automatic assignment,
    # always replaces that placeholder's current mapping.
    # ---------------------------

    def _preview_placeholder_at(self, pos: QtCore.QPoint, allow_fallback=False):
        doc = self.preview.document()
        best_ph = None
        best_dist = float('inf')
        preview_plain_text = self.preview.toPlainText()

        for ph in getattr(self, "placeholders", []) or []:
            visible = self.mapping.get(ph, f"{{{ph}}}")
            if not visible:
                visible = f"{{{ph}}}"
            search_text = strip_markup_to_plain_ascii(visible)
            search_text = remove_subscript_unicode(search_text)
            if not search_text:
                search_text = visible
            search_variants = [search_text]
            if search_text not in preview_plain_text:
                search_variants.append(remove_subscript_unicode(visible))
            for variant in search_variants:
                if not variant:
                    continue
                start = 0
                while True:
                    start = preview_plain_text.find(variant, start)
                    if start < 0:
                        break
                    cursor = QtGui.QTextCursor(doc)
                    cursor.setPosition(start)
                    cursor.setPosition(start + len(variant), QtGui.QTextCursor.KeepAnchor)
                    start_cursor = QtGui.QTextCursor(cursor)
                    start_cursor.clearSelection()
                    start_cursor.setPosition(start)
                    end_cursor = QtGui.QTextCursor(cursor)
                    end_cursor.clearSelection()
                    end_cursor.setPosition(max(start, start + len(variant) - 1))
                    rect = self.preview.cursorRect(start_cursor).united(self.preview.cursorRect(end_cursor))
                    zone = rect.adjusted(-12, -8, 12, 8)
                    if zone.contains(pos):
                        dist = 0
                    else:
                        center = zone.center()
                        dx = center.x() - pos.x()
                        dy = center.y() - pos.y()
                        dist = math.hypot(dx, dy)
                    if dist < best_dist:
                        best_dist = dist
                        best_ph = ph
                    start += max(1, len(variant))

        if not best_ph and allow_fallback:
            all_phs = extract_placeholders(self.current_template)
            if all_phs:
                best_ph = all_phs[0]
        return best_ph

    def on_preview_drag_hovered(self, text: str, pos: QtCore.QPoint):
        if not getattr(self, "current_template", "") or not text.strip():
            return
        # Always resolve against the real mapping so a temporary preview from
        # the previous position cannot lock the drag to its first target.
        self.update_preview_from_markup(self._render_markup(self.mapping))
        self._drag_preview_placeholder = self._preview_placeholder_at(pos)
        placeholder = self._drag_preview_placeholder
        if not placeholder or "──" in text:
            return
        preview_mapping = dict(self.mapping)
        preview_mapping[placeholder] = text.strip()
        self.update_preview_from_markup(
            self._render_markup(preview_mapping),
            highlight_mapping=preview_mapping,
        )

    def clear_preview_drag_hover(self):
        if getattr(self, "_drag_preview_placeholder", None):
            self._drag_preview_placeholder = None
            self.render_preview()

    def on_preview_dropped(self, text: str, pos: QtCore.QPoint):
        if not getattr(self, "current_template", ""):
            QtWidgets.QMessageBox.information(self, "No template", "No procedure/template loaded.")
            return

        best_ph = getattr(self, "_drag_preview_placeholder", None) or self._preview_placeholder_at(pos)
        self._drag_preview_placeholder = None

        if not best_ph:
            QtWidgets.QMessageBox.information(self, "No placeholders", "No placeholder was found in the template to assign to.")
            return

        val = text.strip()
        if val:
            if "──" in val:
                return
            # push state for undo
            self.push_state()
            previous_value = self.mapping.get(best_ph)
            if val == "<unassigned>":
                if best_ph in self.mapping:
                    del self.mapping[best_ph]
            else:
                self.mapping[best_ph] = val
                self.automatic_mapping.discard(best_ph)
            self.render_preview()
            self.refresh_placeholders_list()
            action = "Reassigned" if previous_value else "Assigned"
            self.statusBar().showMessage(f"{action} '{val}' to placeholder {{{best_ph}}}", 3000)

def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setStyleSheet(DARK_STYLESHEET)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()