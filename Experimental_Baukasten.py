# -*- coding: utf-8 -*-
"""
author: vinschu
"""
import sys
import json
import re
import time
from pathlib import Path
from PySide6 import QtWidgets, QtGui, QtCore
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt

import xml.etree.ElementTree as ET
import os
import html as html_std
import urllib.request
import urllib.parse
import urllib.error
import socket

# Standard-Datei-Pfade 
if getattr(sys, 'frozen', False):
    # Wenn als EXE ausgeführt
    APP_DIR = Path(sys.executable).parent
else:
    # Wenn als .py ausgeführt
    APP_DIR = Path(__file__).parent.resolve()

PROCEDURES_FILE = APP_DIR / "procedures.json"
SOLVENTS_FILE = APP_DIR / "solvents.json"
TRIVIALS_FILE = APP_DIR / "trivial_names.json"


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

def extract_placeholders(template_text):
    return list(dict.fromkeys(re.findall(r"\{([a-zA-Z0-9_]+)\}", template_text)))


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

# Reverse mapping: Unicode subscript digits back to ASCII digits
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
# Build a translation table for str.translate
SUBSCRIPT_REVERSE_TRANSLATION = str.maketrans(_UNICODE_SUB_TO_ASCII)
# Set for quick membership testing
UNICODE_SUBSCRIPT_DIGITS = set(_UNICODE_SUB_TO_ASCII.keys())


def subscript_digits(s: str) -> str:
    """
    Replace ASCII digits in s with Unicode subscript digits.
    Only digits are transformed; letters, punctuation and spaces preserved.
    Example: "C18H20" -> "C₁₈H₂₀"
    """
    if not s:
        return s
    return s.translate(SUBSCRIPT_TRANSLATION)


def remove_subscript_unicode(s: str) -> str:
    """
    Convert Unicode subscript digits back to ASCII digits.
    Example: "C₁₈H₂₀" -> "C18H20"
    """
    if not s:
        return s
    return s.translate(SUBSCRIPT_REVERSE_TRANSLATION)


def normalize_commas_in_number_string(s: str) -> str:
    """
    Replace comma decimal separators with dots, but only where comma is between digits.
    Examples:
      '681,80 mg' -> '681.80 mg'
      '2,06mmol'  -> '2.06mmol'
      'C18H20BrN' -> unchanged
    """
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
    """
    Legacy helper that converts digit-markup to unicode digits (kept for compatibility).
    """
    if not s:
        return s
    # convert HTML-like <sub>123</sub> (case-insensitive)
    s = re.sub(r'(?i)<sub>\s*([0-9]+)\s*</sub>',
               lambda m: m.group(1).translate(SUBSCRIPT_TRANSLATION),
               s)
    # convert _{123}
    s = re.sub(r'_\{\s*([0-9]+)\s*\}',
               lambda m: m.group(1).translate(SUBSCRIPT_TRANSLATION),
               s)
    # convert _123
    s = re.sub(r'_([0-9]+)',
               lambda m: m.group(1).translate(SUBSCRIPT_TRANSLATION),
               s)
    return s


# New: render markup to HTML for preview and copy. Handles:
#   _x  and _{...}  -> subscript (wrap in <sub>...</sub>)
#   /x  and /{...}  -> italic (wrap in <i>...</i>)
#   existing <sub>...</sub> is preserved (case-insensitive)
_MARKUP_RE = re.compile(
    r'(?P<htmlsub><sub>.*?</sub>)'  # already HTML sub tag (non-greedy)
    r'|_\{\s*([^}]+)\s*\}'          # _{...}  (group 2)
    r'|_([A-Za-z0-9])'              # _x (group 3)
    r'|/\{\s*([^}]+)\s*\}'          # /{...}  (group 4)
    r'|/([A-Za-z0-9])',             # /x (group 5)
    flags=re.IGNORECASE | re.DOTALL
)


def render_markup_to_html(s: str) -> str:
    """
    Convert markup in s to HTML. Returns HTML fragment (no <html> wrapper).
    Examples:
      H_2 -> H<sub>2</sub>
      C_{18} -> C<sub>18</sub>
      _f -> <sub>f</sub>
      /R -> <i>R</i>
      /{R} -> <i>R</i>
      <sub>2</sub> (if present) preserved
    """
    if s is None:
        return ""
    out = []
    last = 0
    for m in _MARKUP_RE.finditer(s):
        start, end = m.span()
        if start > last:
            out.append(html_std.escape(s[last:start]))
        # groups: m.group(1) is htmlsub (if matched) else group2/group3/group4/group5
        htmlsub = m.group(1)
        brace_sub = m.group(2)
        single_sub = m.group(3)
        brace_i = m.group(4)
        single_i = m.group(5)
        if htmlsub:
            # preserve inner content escaped
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
    """
    Produce a plain-ASCII version by removing the markup characters.
    _x -> x, _{...} -> ..., /x -> x, /{...} -> ...
    Also strips any HTML <sub>/<i> tags if present in the literal text.
    """
    if s is None:
        return ""
    # replace <sub>..</sub> with inner
    s = re.sub(r'(?i)<sub>\s*([^<]+?)\s*</sub>', r'\1', s)
    # replace <i>..</i>
    s = re.sub(r'(?i)<i>\s*([^<]+?)\s*</i>', r'\1', s)
    # _{...} -> inner
    s = re.sub(r'_\{\s*([^}]+)\s*\}', r'\1', s)
    # _x -> x
    s = re.sub(r'_([A-Za-z0-9])', r'\1', s)
    # /{...} -> inner
    s = re.sub(r'/\{\s*([^}]+)\s*\}', r'\1', s)
    # /x -> x
    s = re.sub(r'/([A-Za-z0-9])', r'\1', s)
    return s


# ---------------------------
# CDXML / stoichiometry parsing (component-wise)
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


def parse_stoichiometry_components(cdxml_path):
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

    # exclude header component, then ignore the first non-header split
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
            if fname.lower().endswith((".cdxml", ".xml")):
                full = os.path.join(dirpath, fname)
                try:
                    with open(full, "rb") as fh:
                        head = fh.read(400).lower()
                    if b"cdxml" in head or b"chem" in head or fname.lower().endswith(".cdxml"):
                        cdxml_files.append(full)
                except Exception:
                    if fname.lower().endswith(".cdxml"):
                        cdxml_files.append(full)
    return cdxml_files


def format_component_display_by_index(component, is_last=False):
    vals = component.get("values", [])

    def safe(i):
        try:
            return (vals[i] or "").strip()
        except Exception:
            return ""

    row0 = safe(0)
    row4 = safe(4)
    row5 = safe(5)
    row9 = safe(9)
    row10 = safe(10)

    # normalize numeric commas to dots where appropriate
    row0_disp = normalize_commas_in_number_string(row0)
    row4_disp = normalize_commas_in_number_string(row4)
    row5_disp = normalize_commas_in_number_string(row5)
    row9_disp = normalize_commas_in_number_string(row9)
    row10_disp = normalize_commas_in_number_string(row10)

    # append " eq." to value[4] for display (if present)
    row4_eq_disp = append_eq_if_present(row4_disp) if row4_disp else ""

    # choose value[9] if not zero-ml else fallback to value[5]
    chosen = row9_disp if (row9 and not is_zero_ml(row9)) else row5_disp

    # formula handling: raw formula (for dedupe/ident) and subscripted formula for display
    formula_raw = component.get("formula", "").strip() or row0
    formula_raw = re.sub(r'^[\s\-_\.0-9:()]+', '', formula_raw).strip()
    formula_display = subscript_digits(formula_raw)

    if is_last:
        # last split: only show the formula (no bracket)
        return {"formula": formula_raw, "display": formula_display, "details": {"row0": row0_disp}}

    # Normal case: bracket order -> chosen (row9 or 5), row10, row4_eq
    bracket_parts = []
    for part in (chosen, row10_disp, row4_eq_disp):
        if part:
            bracket_parts.append(part)
    bracket = ", ".join(bracket_parts)

    if bracket:
        display = f"{formula_display} ({bracket})"
    else:
        display = formula_display

    return {"formula": formula_raw, "display": display, "details": {"chosen": chosen, "row10": row10_disp, "row4_eq": row4_eq_disp}}


# ---------------------------
# Solvent editor & manager dialogs
# (unchanged)
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
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(del_btn)
        layout.addLayout(btn_layout)

        add_btn.clicked.connect(self.add_solvent)
        edit_btn.clicked.connect(self.edit_solvent)
        del_btn.clicked.connect(self.delete_solvent)

        close_btn = QtWidgets.QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)

    def refresh_list(self):
        self.list_widget.clear()
        for s in self.solvents:
            # show subscripted in manager view for clarity
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


# ---------------------------
# Procedure editor & manager dialogs
# (unchanged)
# ---------------------------

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
    """
    Manage procedures (add / edit / delete). Uses ProcedureEditorDialog for editing.
    The dialog exposes `self.procedures` on accept (list of procedure dicts).
    """
    def __init__(self, parent=None, procedures=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Procedures")
        self.resize(800, 450)
        # store a shallow copy so edits here don't immediately modify caller's list
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
        # spacer
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
            # generate a simple unique id using timestamp
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

# ---------------------------
# New: Trivial Names manager dialog
# ---------------------------

class TrivialNamesManagerDialog(QtWidgets.QDialog):
    """
    Manage mapping from ASCII formula -> trivial name.
    """
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
        # show keys with subscripted digits for readability
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
                # remove old key if changed
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



# ---------------------------
# Edit Output dialog (unchanged)
# ---------------------------

class EditOutputDialog(QtWidgets.QDialog):
    """
    Dialog that allows the user to edit the rendered markup.
    The content edited here is stored temporarily (not written into the procedure templates).
    """
    def __init__(self, parent=None, initial_markup=""):
        super().__init__(parent)
        self.setWindowTitle("Edit Output (temporary)")
        self.resize(700, 400)
        layout = QtWidgets.QVBoxLayout(self)

        self.editor = QtWidgets.QPlainTextEdit()
        # show the raw markup so user can edit subscripts (/ and _ not rendered)
        self.editor.setPlainText(initial_markup or "")
        layout.addWidget(self.editor)

        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def get_text(self):
        return self.editor.toPlainText()


# ---------------------------
# Drop area for files/folders (main window)
# ---------------------------

class DropArea(QtWidgets.QWidget):
    """
    Accepts file/dir drops. Emits signal with list of local paths.
    """
    filesDropped = QtCore.Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        # optional visual hint
        self.setAutoFillBackground(True)
        pal = self.palette()
        # keep default background; visual hint not intrusive
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



# ---------------------------
# Main window
# ---------------------------

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Experimental Section Writer")
        self.resize(1000, 700)
        self.procedures = load_procedures()
        self.solvents = load_solvents()
        self.trivial_names = load_trivial_names()
        self.chemicals = []
        self.mapping = {}
        self.last_rendered_markup = ""

        self.setup_ui()

    def setup_ui(self):
        # Menu
        men = self.menuBar()
        file_menu = men.addMenu("File")
        load_act = QtGui.QAction("Load CDMXL...", self)
        load_act.triggered.connect(self.load_cdmxl)
        file_menu.addAction(load_act)

        procedures_menu = men.addMenu("Procedures")
        manage_act = QtGui.QAction("Manage Procedures...", self)
        manage_act.triggered.connect(self.open_procedure_manager)
        procedures_menu.addAction(manage_act)

        # Add a top-level "Solvents" action to the menu bar (opens solvent manager)
        solvents_act = QtGui.QAction("Solvents", self)
        solvents_act.triggered.connect(self.open_solvent_manager)
        men.addAction(solvents_act)

        # New: Names menu for trivial names management
        names_menu = men.addMenu("Names")
        manage_triv_act = QtGui.QAction("Manage Trivial Names...", self)
        manage_triv_act.triggered.connect(self.open_trivial_names_manager)
        names_menu.addAction(manage_triv_act)

        # Central widget: use DropArea so user can drop files/folders onto entire main area
        central = DropArea(self)
        central.filesDropped.connect(self.load_cdxml_paths)
        self.setCentralWidget(central)
        layout = QtWidgets.QHBoxLayout(central)

        left = QtWidgets.QVBoxLayout()
        right = QtWidgets.QVBoxLayout()

        # Chemicals list
        left.addWidget(QtWidgets.QLabel("Chemicals (from CDMXL)"))
        self.chem_list = QtWidgets.QListWidget()
        left.addWidget(self.chem_list)

        # Solvents list under chemicals
        left.addWidget(QtWidgets.QLabel("Solvents"))
        solv_h = QtWidgets.QHBoxLayout()
        self.solv_list = QtWidgets.QListWidget()
        solv_h.addWidget(self.solv_list, 1)
        left.addLayout(solv_h)
        self.refresh_solvents_list()

        # Right side: procedures / placeholders / preview
        right_top = QtWidgets.QVBoxLayout()
        proc_h = QtWidgets.QHBoxLayout()
        proc_h.addWidget(QtWidgets.QLabel("Procedure:"))
        self.proc_combo = QtWidgets.QComboBox()
        proc_h.addWidget(self.proc_combo)

        # ====== RESTORED BUTTONS (Load) placed next to procedure combo ======
        # These are visible UI buttons (in addition to menu actions) requested by the user.
        # They call the same methods that the menu/actions use, so behavior is unchanged.
        spacer = QtWidgets.QWidget()
        spacer.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        proc_h.addWidget(spacer)

        load_btn = QtWidgets.QPushButton("Load")
        load_btn.setToolTip("Load CDMXL files or folders (same as File → Load CDMXL...)")
        load_btn.clicked.connect(self.load_cdmxl)
        proc_h.addWidget(load_btn)

       
        # =======================================================================================

        right_top.addLayout(proc_h)

        # Create placeholders and mapping widgets first to ensure on_proc_changed can safely access mapping_list
        right_top.addWidget(QtWidgets.QLabel("Placeholders (click to assign)"))
        self.placeholders_list = QtWidgets.QListWidget()
        self.placeholders_list.itemClicked.connect(self.on_placeholder_clicked)
        right_top.addWidget(self.placeholders_list)

        right_top.addWidget(QtWidgets.QLabel("Assigned mapping"))
        self.mapping_list = QtWidgets.QListWidget()
        right_top.addWidget(self.mapping_list)

        # Now connect the combo and load procedures (safe because placeholders/mapping exist)
        self.proc_combo.currentIndexChanged.connect(self.on_proc_changed)
        self.reload_procedures()

        right.addLayout(right_top)

        # Preview and export
        preview_lbl = QtWidgets.QLabel("Preview")
        right.addWidget(preview_lbl)
        self.preview = QtWidgets.QTextEdit()
        self.preview.setReadOnly(True)
        right.addWidget(self.preview, 1)

        bottom_h = QtWidgets.QHBoxLayout()
        # Keep the bottom buttons too (they were present previously). They still call the same methods:
        preview_btn = QtWidgets.QPushButton("Render Preview")
        preview_btn.clicked.connect(self.render_preview)
        bottom_h.addWidget(preview_btn)

        copy_preview_btn = QtWidgets.QPushButton("Copy Preview")
        copy_preview_btn.clicked.connect(self.copy_preview_to_clipboard)
        bottom_h.addWidget(copy_preview_btn)

        export_word_btn = QtWidgets.QPushButton("Export to Word")
        export_word_btn.clicked.connect(self.export_word)
        bottom_h.addWidget(export_word_btn)

        # New: Edit Output button (temporary edit of last rendered markup)
        edit_output_btn = QtWidgets.QPushButton("Edit Output")
        edit_output_btn.setToolTip("Edit the generated output text (temporary). Save to update the preview; changes are not written to procedures.")
        edit_output_btn.clicked.connect(self.open_edit_output_dialog)
        bottom_h.addWidget(edit_output_btn)

        # NEW: Convert Names button
        convert_btn = QtWidgets.QPushButton("Convert Names")
        convert_btn.setToolTip("Convert chemical formulas in the preview to trivial names (from file) or IUPAC (PubChem fallback).")
        convert_btn.clicked.connect(self.convert_preview_formulas_to_names)
        bottom_h.addWidget(convert_btn)

        right.addLayout(bottom_h)

        layout.addLayout(left, 1)
        layout.addLayout(right, 2)

        status = QtWidgets.QLabel("")
        self.statusBar().addWidget(status)

    # ---------------------------
    # Loading CDXML files (from dialog or drag/drop) - centralised
    # ---------------------------

    def load_cdmxl(self):
        fpath, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Open CDMXL file (or cancel to choose folder)", str(Path.cwd()),
            "CDMXL Files (*.cdxml *.xml);;All files (*)"
        )
        if fpath:
            # single file chosen
            self.load_cdxml_paths([fpath])
            return

        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select folder to scan for CDXML files", str(Path.cwd()))
        if not folder:
            return
        self.load_cdxml_paths([folder])

    def load_cdxml_paths(self, paths):
        """
        Accept a list of filesystem paths (files or directories). Resolve to a list of files
        to parse, then parse and populate the chemicals list exactly like the previous dialog flow.
        """
        to_parse = []
        for p in paths:
            if os.path.isfile(p):
                to_parse.append(p)
            elif os.path.isdir(p):
                found = find_cdxml_files(p)
                if found:
                    to_parse.extend(found)
        # dedupe
        to_parse = list(dict.fromkeys(to_parse))
        if not to_parse:
            QtWidgets.QMessageBox.information(self, "No files", "No CDXML/XML files found in the dropped items.")
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
                "formula": formatted["formula"],  # raw for dedupe
                "display": formatted["display"]   # subscripted display
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
            return

        self.chemicals = normalized
        self.chem_list.clear()
        for c in self.chemicals:
            self.chem_list.addItem(c["display"])
        # update status bar (silent confirmation)
        self.statusBar().showMessage(f"Loaded {len(self.chemicals)} candidate chemicals", 5000)

    # ---------------------------
    # Rest of main window methods (placeholders, preview, copy/export etc.)
    # ---------------------------

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
        if hasattr(self, "placeholders_list") and hasattr(self, "mapping_list") and self.procedures:
            self.proc_combo.setCurrentIndex(0)
            self.on_proc_changed(0)

    def open_procedure_manager(self):
        dlg = ProcedureManagerDialog(self, procedures=self.procedures)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            self.procedures = dlg.procedures
            save_procedures(self.procedures)
            self.reload_procedures()

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
        if hasattr(self, "placeholders_list"):
            self.placeholders_list.clear()
            for ph in phs:
                item = QtWidgets.QListWidgetItem(ph)
                self.placeholders_list.addItem(item)
        if hasattr(self, "mapping_list"):
            self.refresh_mapping_list()
        if hasattr(self, "preview"):
            self.preview.clear()
            self.last_rendered_markup = ""

    def on_placeholder_clicked(self, item):
        placeholder = item.text()

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

            if text == "<unassigned>":
                if placeholder in self.mapping:
                    del self.mapping[placeholder]
            elif "──" not in text:  # ignore separators
                self.mapping[placeholder] = text

            self.refresh_mapping_list()
            self.render_preview()
            dlg.accept()  # close instantly

        listw.itemClicked.connect(on_item_clicked)

        dlg.exec()



    def refresh_mapping_list(self):
        if not hasattr(self, "mapping_list"):
            return
        self.mapping_list.clear()
        for ph in self.placeholders:
            val = self.mapping.get(ph, "")
            li = QtWidgets.QListWidgetItem(f"{ph} → {val or '<unassigned>'}")
            self.mapping_list.addItem(li)

    def render_preview(self):
        """
        Build rendered text by replacing placeholders with assigned values,
        store that markup string in self.last_rendered_markup, then render HTML in preview
        so subscripts (including letter subscripts) and italics show visually.
        """
        if not hasattr(self, "current_template"):
            return
        text = self.current_template
        def repl(m):
            key = m.group(1)
            return self.mapping.get(key, f"{{{key}}}")
        rendered = re.sub(r"\{([a-zA-Z0-9_]+)\}", repl, text)
        self.last_rendered_markup = rendered
        self.update_preview_from_markup(rendered)

    def update_preview_from_markup(self, markup):
        """
        Given the markup (raw text with _{..}, /{..}, etc.), render HTML and set to preview widget.
        """
        paragraphs = []
        for line in markup.splitlines() or [""]:
            frag = render_markup_to_html(line)
            paragraphs.append(f'<p style="text-align:justify; line-height:1.5; margin:0 0 6px 0;">{frag}</p>')
        html_full = "<html><body>" + "".join(paragraphs) + "</body></html>"
        # set HTML to preview
        if hasattr(self, "preview"):
            self.preview.setHtml(html_full)

    def open_edit_output_dialog(self):
        """
        Open dialog to edit the current last_rendered_markup.
        If there's no last_rendered_markup, try to generate via render_preview() first.
        Saving in the dialog updates self.last_rendered_markup and preview, but does NOT write
        changes to the procedure/template file — it's temporary only.
        """
        if not getattr(self, "last_rendered_markup", "").strip():
            # ensure something is generated
            self.render_preview()
        initial = getattr(self, "last_rendered_markup", "") or ""
        dlg = EditOutputDialog(self, initial_markup=initial)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            new_text = dlg.get_text() or ""
            # store edited markup temporarily
            self.last_rendered_markup = new_text
            self.update_preview_from_markup(new_text)

    def _convert_unicode_subs_to_html_subs(self, s: str) -> str:
        """
        Convert sequences of Unicode subscript digits (₀₁₂...) into HTML <sub>...</sub>
        where the digits inside are the ASCII equivalents. This ensures the HTML clipboard
        contains <sub> tags (Word will import them as formatted subscript runs),
        instead of plain Unicode subscript characters.
        Example: "Et₂O" -> "Et<sub>2</sub>O"
        """
        if not s:
            return s
        # build a character class for the known unicode subscript digits
        chars = "".join(re.escape(ch) for ch in UNICODE_SUBSCRIPT_DIGITS)
        pattern = re.compile(f"[{chars}]+")
        def repl(m):
            uni = m.group(0)
            ascii_digits = uni.translate(SUBSCRIPT_REVERSE_TRANSLATION)
            return f"<sub>{ascii_digits}</sub>"
        return pattern.sub(repl, s)

    def copy_preview_to_clipboard(self):
        """
        Copy the last rendered markup to clipboard as HTML and plain text.
        HTML uses <sub> and <i> tags and paragraph styling (justify + 1.5) so Word will paste exact formatting.
        Important: mapping/preview may show Unicode subscript digits for on-screen display.
        Here we convert any Unicode subscript digits back to HTML <sub>..</sub> so Word receives
        formatted subs (not raw Unicode characters) and the pasted result matches Word formatting.
        """
        markup = getattr(self, "last_rendered_markup", "") or ""
        if not markup.strip():
            return

        # Convert Unicode subscript digits (from subscript_digits earlier) into <sub>digit</sub>
        markup_for_html = self._convert_unicode_subs_to_html_subs(markup)

        paragraphs = []
        for line in markup_for_html.splitlines() or [""]:
            frag = render_markup_to_html(line)
            paragraphs.append(f'<p style="text-align:justify; line-height:1.5; margin:0 0 6px 0;">{frag}</p>')
        html_full = f"<html><body>{''.join(paragraphs)}</body></html>"

        # Plain ASCII: remove markup and convert unicode subs back to ASCII digits
        plain_lines = []
        for line in markup.splitlines() or [""]:
            # first convert unicode subscript digits to ascii so plain text contains normal digits
            ascii_line = remove_subscript_unicode(line)
            ascii_line = strip_markup_to_plain_ascii(ascii_line)
            plain_lines.append(ascii_line)
        plain_ascii = "\n".join(plain_lines)

        mime = QtCore.QMimeData()
        mime.setHtml(html_full)
        mime.setText(plain_ascii)
        cb = QtWidgets.QApplication.clipboard()
        cb.setMimeData(mime)
        # silent copy

    def export_word(self):
        """
        Export the last rendered markup to a Word document.
        Applies paragraph formatting: justified alignment and 1.5 line spacing.
        Converts markup to subscript/italic runs.
        """
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
            r'|/(?P<single_i>[A-Za-z0-9])',
            flags=re.IGNORECASE | re.DOTALL
        )
        try:
            doc = Document()
            for line in markup.splitlines() or [""]:
                p = doc.add_paragraph()
                # set paragraph formatting: justified, 1.5 line spacing
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
                    if m.group('htmlsub'):
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

    # ---------------------------
    # Trivial names management UI wrapper
    # ---------------------------

    def open_trivial_names_manager(self):
        dlg = TrivialNamesManagerDialog(self, mapping=self.trivial_names)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            self.trivial_names = dlg.mapping
            save_trivial_names(self.trivial_names)
            self.statusBar().showMessage(f"Saved {len(self.trivial_names)} trivial name entries", 4000)

    # ---------------------------
    # Convert names in preview
    # ---------------------------

    def convert_preview_formulas_to_names(self):
        """
        Convert formulas in self.last_rendered_markup to trivial names (from file) or IUPAC (PubChem fallback).
        Updates self.last_rendered_markup and the preview.
        """
        markup = getattr(self, "last_rendered_markup", "") or ""
        if not markup.strip():
            QtWidgets.QMessageBox.information(self, "Nothing", "No rendered text to convert.")
            return

        # Build trivial mapping (keys normalized to ASCII, uppercase sensitive)
        trivial_map = {k: v for k, v in self.trivial_names.items()}

        # find candidate tokens in markup that look like formulas:
        # token chars: letters, digits, unicode subscript digits, parentheses
        token_candidates = set(re.findall(r'[A-Za-z₀₁₂₃₄₅₆₇₈₉0-9\(\)]+', markup))
        # normalize candidates and filter those that contain at least one letter and at least one digit (likely formulas)
        candidates = []
        for tok in token_candidates:
            plain = remove_subscript_unicode(strip_markup_to_plain_ascii(tok)).strip()
            if not plain:
                continue
            if re.search(r'[A-Za-z]', plain) and re.search(r'\d', plain):
                candidates.append((tok, plain))

        if not candidates:
            QtWidgets.QMessageBox.information(self, "No formulas", "No formula-like tokens found to convert.")
            return

        # Sort candidates by length of plain formula descending (avoid partial replacements)
        candidates.sort(key=lambda x: len(x[1]), reverse=True)

        new_markup = markup
        looked_up = {}
        pubchem_cache = {}

        for orig_tok, plain in candidates:
            plain_norm = plain.replace(" ", "")
            # first: try trivial mapping (exact ASCII key)
            trivial_name = trivial_map.get(plain_norm)
            if trivial_name:
                # replace occurrences of the original token (which may have unicode subs) with trivial name
                # Use simple string replace for the exact token text to preserve other markup
                new_markup = new_markup.replace(orig_tok, trivial_name)
                looked_up[plain_norm] = ("trivial", trivial_name)
                continue

            # second: try matching ascii form of token (maybe markup removed)
            # also try replacing any unicode-subscript variant
            unicode_variant = subscript_digits(plain_norm)
            replaced = False
            if unicode_variant in new_markup:
                # Double-check trivial mapping for ascii (already done), so fallback to PubChem
                # We'll attempt PubChem lookup for plain_norm
                pass

            # If not in trivial map, attempt PubChem (best-effort)
            if plain_norm in pubchem_cache:
                iupac = pubchem_cache[plain_norm]
            else:
                # Attempt network lookup (may be slow). Do not raise on error.
                iupac = lookup_iupac_from_pubchem_by_formula(plain_norm)
                pubchem_cache[plain_norm] = iupac
                # be a bit gentle on PubChem if many queries
                time.sleep(0.15)

            if iupac:
                # Replace occurrences of orig_tok (token as in markup) and also ascii and unicode variants
                new_markup = new_markup.replace(orig_tok, iupac)
                # also replace raw ASCII occurrences (rare) and unicode variant
                new_markup = new_markup.replace(plain_norm, iupac)
                new_markup = new_markup.replace(unicode_variant, iupac)
                looked_up[plain_norm] = ("iupac", iupac)
            else:
                looked_up[plain_norm] = ("notfound", None)

        # If there was any successful replacement, update
        any_changed = new_markup != markup
        if any_changed:
            self.last_rendered_markup = new_markup
            self.update_preview_from_markup(new_markup)
            # summary message
            found_count = sum(1 for v in looked_up.values() if v[0] in ("trivial", "iupac"))
            self.statusBar().showMessage(f"Converted {found_count} formula(s) (trivial/iupac) in preview", 6000)
            QtWidgets.QMessageBox.information(self, "Converted", f"Converted {found_count} formula(s) in the preview.")
        else:
            QtWidgets.QMessageBox.information(self, "No changes", "No conversions were possible (no trivial names found and no IUPAC names from PubChem).")

def main():
    app = QtWidgets.QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

