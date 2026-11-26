"""엑셀 파일의 첫 번째 시트를 목록으로 보여주고 행을 관리(CRUD)하는 프로그램.

- 입력, 조회, 수정은 별도의 폼에서 하고, 삭제는 확인을 받은 뒤 실행한다.
- 값이 2개 이상 있는 첫 행을 제목 행으로 보고, 그 아래 행만 수정/삭제한다.
- 변경 내용은 엑셀 파일에 바로 저장하며, 처음 변경하기 전에 백업본을 만든다.
- 마지막으로 선택한 파일은 manager_config.json에 저장해 다음 실행 때 자동으로 연다.
- 추가 설치 없이 파이썬 기본 라이브러리만 사용한다. (.xlsx, .xlsm 지원)
"""

import json
import os
import posixpath
import re
import shutil
import tempfile
import tkinter as tk
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from xml.dom import minidom
from xml.etree import ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"m": MAIN_NS}
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
MAX_ROWS = 10000
# 마지막으로 선택한 파일 경로를 저장하는 설정 파일 (이 프로그램과 같은 폴더)
CONFIG_FILE = Path(__file__).resolve().parent / "manager_config.json"

# 엑셀에 내장된 날짜/시간 형식 번호 (27 이후 일부는 한국어 Excel용)
DATE_IDS = {14, 15, 16, 17, 22, 27, 28, 29, 30, 31, 36, 50, 51, 52, 53, 54, 57, 58}
TIME_IDS = {18, 19, 20, 21, 45, 46, 47}


class ExcelError(Exception):
    """사용자에게 그대로 보여 줄 수 있는 오류."""


# ---------------------------------------------------------------- 엑셀 읽기

def check_signature(path):
    with open(path, "rb") as f:
        head = f.read(16)
    if head.startswith(b"PK"):
        return
    if head.startswith(b"BMS DocuRay"):
        raise ExcelError(
            "문서보안(DRM)으로 암호화된 파일이라 읽을 수 없습니다.\n"
            "사내 절차에 따라 보안이 해제된 파일을 선택해 주세요."
        )
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        raise ExcelError(
            "암호가 걸렸거나 예전 형식(.xls)인 파일입니다.\n"
            "암호 없이 .xlsx로 저장한 뒤 선택해 주세요."
        )
    raise ExcelError("올바른 .xlsx 파일이 아니거나 암호화된 파일입니다.")


def read_xml(archive, name):
    """압축 파일 안의 XML을 읽는다. 없으면 None."""
    try:
        return ET.fromstring(archive.read(name))
    except KeyError:
        return None


def text_of(node):
    """서식이 나뉜 문자열을 합친다. 윗주(rPh)는 제외한다."""
    parts = node.findall("m:t", NS) + node.findall("m:r/m:t", NS)
    return "".join(t.text or "" for t in parts)


def read_styles(archive):
    """셀 스타일 번호 순서대로 (형식 번호, 형식 문자열) 목록을 만든다."""
    root = read_xml(archive, "xl/styles.xml")
    if root is None:
        return []
    custom = {
        f.get("numFmtId"): f.get("formatCode")
        for f in root.findall("m:numFmts/m:numFmt", NS)
    }
    return [
        (int(xf.get("numFmtId", 0)), custom.get(xf.get("numFmtId"), ""))
        for xf in root.findall("m:cellXfs/m:xf", NS)
    ]


def analyze_style(style):
    """(정리한 형식 문자열, 날짜 여부, 시간 여부, 퍼센트 여부)"""
    format_id, code = style
    # 따옴표 문자, [색상] 등을 제거하고 양수용 형식만 확인
    code = re.sub(r'"[^"]*"|\[[^\]]*\]|[\\_*].', "", code).split(";")[0].lower()
    is_date = format_id in DATE_IDS or bool(re.search(r"[yd]", code))
    is_time = format_id in TIME_IDS or bool(re.search(r"[hs]", code))
    is_percent = format_id in (9, 10) or "%" in code
    return code, is_date, is_time, is_percent


def date_base(date1904):
    return datetime(1904, 1, 1) if date1904 else datetime(1899, 12, 30)


def format_number(raw, style, date1904):
    value = float(raw)
    format_id = style[0]
    code, is_date, is_time, is_percent = analyze_style(style)

    if (is_date or is_time) and 0 <= value < 2958466:
        moment = date_base(date1904) + timedelta(seconds=round(value * 86400))
        if is_date and (is_time or format_id == 22):
            return moment.strftime("%Y-%m-%d %H:%M:%S")
        return moment.strftime("%Y-%m-%d" if is_date else "%H:%M:%S")

    decimals_match = re.search(r"\.(0+)", code)
    decimals = len(decimals_match.group(1)) if decimals_match else 0
    if is_percent:
        if format_id == 10:
            decimals = 2
        return f"{value * 100:.{decimals}f}%"
    if format_id in (3, 4) or 37 <= format_id <= 44 or "#,##" in code:
        if format_id in (4, 39, 40, 43, 44):
            decimals = 2
        return f"{value:,.{decimals}f}"
    return f"{value:.15g}"


def column_number(ref):
    """'AB12' -> 28"""
    number = 0
    for ch in re.match(r"[A-Z]+", ref).group():
        number = number * 26 + ord(ch) - ord("A") + 1
    return number


def row_of(ref):
    """'AB12' -> 12"""
    return int(re.search(r"\d+", ref).group())


def column_name(number):
    """28 -> 'AB'"""
    name = ""
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        name = chr(ord("A") + remainder) + name
    return name


def locate_first_sheet(archive):
    """(workbook XML, 첫 번째 시트 이름, 시트 XML의 압축 파일 내 경로)"""
    workbook = read_xml(archive, "xl/workbook.xml")
    rels = read_xml(archive, "xl/_rels/workbook.xml.rels")
    sheet = None if workbook is None else workbook.find("m:sheets/m:sheet", NS)
    if sheet is None or rels is None:
        raise ExcelError("엑셀 통합 문서에서 시트를 찾을 수 없습니다.")

    target = next(r.get("Target") for r in rels if r.get("Id") == sheet.get(REL_ID))
    if target.startswith("/"):
        sheet_path = target[1:]
    else:
        sheet_path = posixpath.normpath("xl/" + target)
    if sheet_path not in archive.namelist():
        raise ExcelError("첫 번째 시트의 내용을 찾을 수 없습니다.")
    return workbook, sheet.get("name"), sheet_path


def is_date1904(workbook):
    properties = workbook.find("m:workbookPr", NS)
    return properties is not None and properties.get("date1904") in ("1", "true")


def read_first_sheet(path):
    """(시트 이름, [(행 번호, {열 번호: 값})])를 돌려준다."""
    check_signature(path)
    with zipfile.ZipFile(path) as archive:
        workbook, name, sheet_path = locate_first_sheet(archive)
        sheet_xml = read_xml(archive, sheet_path)
        shared = read_xml(archive, "xl/sharedStrings.xml")
        strings = [] if shared is None else [text_of(si) for si in shared.findall("m:si", NS)]
        styles = read_styles(archive)
        date1904 = is_date1904(workbook)

        rows = []
        row_number = 0
        for row in sheet_xml.findall("m:sheetData/m:row", NS):
            row_number = int(row.get("r", row_number + 1))
            cells = {}
            column = 0
            for cell in row.findall("m:c", NS):
                ref = cell.get("r")
                column = column_number(ref) if ref else column + 1
                kind = cell.get("t", "n")
                raw = cell.findtext("m:v", "", NS)

                if kind == "inlineStr":
                    inline = cell.find("m:is", NS)
                    text = "" if inline is None else text_of(inline)
                elif raw == "":
                    continue
                elif kind == "s":
                    text = strings[int(raw)]
                elif kind == "b":
                    text = "TRUE" if raw == "1" else "FALSE"
                elif kind == "n":
                    index = int(cell.get("s", 0))
                    style = styles[index] if index < len(styles) else (0, "")
                    text = format_number(raw, style, date1904)
                else:
                    text = raw

                if text:
                    cells[column] = text
            if cells:
                rows.append((row_number, cells))

        return name, rows


# ---------------------------------------------------------------- 엑셀 쓰기

def to_number(text, style, date1904):
    """입력한 글자를 숫자로 저장할 수 있으면 그 값, 아니면 None."""
    _, is_date, is_time, is_percent = analyze_style(style)
    value = text.strip()

    if is_date:
        for pattern in ("%Y-%m-%d", "%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S"):
            try:
                delta = datetime.strptime(value, pattern) - date_base(date1904)
            except ValueError:
                continue
            return repr(delta.total_seconds() / 86400)
    if is_time:
        match = re.fullmatch(r"(\d{1,2}):([0-5]\d)(?::([0-5]\d))?", value)
        if match:
            hour, minute, second = (int(g or 0) for g in match.groups())
            return repr((hour * 3600 + minute * 60 + second) / 86400)

    percent = value.endswith("%")
    if percent:
        value = value[:-1].strip()
    # 007처럼 0으로 시작하는 값이나 15자리를 넘는 값은 글자로 둔다.
    if not re.fullmatch(r"-?(0|[1-9]\d{0,2}(,\d{3})+|[1-9]\d*)(\.\d+)?", value):
        return None
    value = value.replace(",", "")
    if len(value.lstrip("-").replace(".", "")) > 15:
        return None
    if percent or is_percent:
        return repr(float(value) / 100)
    return value


class SheetEditor:
    """첫 번째 시트의 XML만 고쳐서 나머지 내용과 서식은 그대로 둔다."""

    def __init__(self, path):
        check_signature(path)
        self.path = path
        with zipfile.ZipFile(path) as archive:
            workbook, _, self.sheet_path = locate_first_sheet(archive)
            self.date1904 = is_date1904(workbook)
            self.styles = read_styles(archive)
            self.document = minidom.parseString(archive.read(self.sheet_path))
        self.drop_calc_chain = False

        sheet_data = self.elements(self.document, "sheetData")
        if not sheet_data:
            raise ExcelError("시트 구조를 해석할 수 없어 수정할 수 없습니다.")
        self.sheet_data = sheet_data[0]

        # 위치 정보(r)가 생략된 행과 셀에 번호를 채운다.
        number = 0
        for row in self.elements(self.sheet_data, "row"):
            number = int(row.getAttribute("r") or number + 1)
            row.setAttribute("r", str(number))
            column = 0
            for cell in self.elements(row, "c"):
                ref = cell.getAttribute("r")
                column = column_number(ref) if ref else column + 1
                cell.setAttribute("r", f"{column_name(column)}{number}")

        self.merges = []  # (첫 행, 첫 열, 끝 행, 끝 열)
        for merge in self.elements(self.document, "mergeCell"):
            start, _, end = merge.getAttribute("ref").partition(":")
            end = end or start
            self.merges.append(
                (row_of(start), column_number(start), row_of(end), column_number(end))
            )

    @staticmethod
    def elements(node, name):
        return node.getElementsByTagNameNS(MAIN_NS, name)

    def create(self, name):
        prefix = self.document.documentElement.prefix
        qualified = f"{prefix}:{name}" if prefix else name
        return self.document.createElementNS(MAIN_NS, qualified)

    def find_row(self, number):
        for row in self.elements(self.sheet_data, "row"):
            if int(row.getAttribute("r")) == number:
                return row
        return None

    def get_row(self, number):
        """행이 없으면 순서에 맞게 새로 만든다."""
        following = None
        for row in self.elements(self.sheet_data, "row"):
            current = int(row.getAttribute("r"))
            if current == number:
                return row
            if current > number:
                following = row
                break
        row = self.create("row")
        row.setAttribute("r", str(number))
        self.sheet_data.insertBefore(row, following)
        return row

    def find_cell(self, row, column):
        if row is None:
            return None
        for cell in self.elements(row, "c"):
            if column_number(cell.getAttribute("r")) == column:
                return cell
        return None

    def get_cell(self, row, column):
        cell = self.find_cell(row, column)
        if cell is not None:
            return cell
        following = next(
            (c for c in self.elements(row, "c")
             if column_number(c.getAttribute("r")) > column),
            None,
        )
        cell = self.create("c")
        cell.setAttribute("r", f"{column_name(column)}{row.getAttribute('r')}")
        row.insertBefore(cell, following)
        return cell

    def locked_reason(self, number, column):
        """입력할 수 없는 칸이면 그 이유, 아니면 None."""
        cell = self.find_cell(self.find_row(number), column)
        if cell is not None and self.elements(cell, "f"):
            return "수식"
        for top, left, bottom, right in self.merges:
            if top <= number <= bottom and left <= column <= right:
                if (number, column) != (top, left):
                    return "병합된 셀"
        return None

    def set_value(self, cell, text, keep_text):
        while cell.firstChild:
            cell.removeChild(cell.firstChild)
        if cell.hasAttribute("t"):
            cell.removeAttribute("t")
        # XML에 저장할 수 없는 제어 문자는 제거
        text = re.sub(r"[^\t\n\r\x20-퟿-�\U00010000-\U0010ffff]", "", text)
        if text == "":
            return

        index = int(cell.getAttribute("s") or 0)
        style = self.styles[index] if index < len(self.styles) else (0, "")
        number = None if keep_text else to_number(text, style, self.date1904)
        if number is not None:
            value = self.create("v")
            value.appendChild(self.document.createTextNode(number))
            cell.appendChild(value)
        else:
            cell.setAttribute("t", "inlineStr")
            inline = self.create("is")
            content = self.create("t")
            content.setAttribute("xml:space", "preserve")
            content.appendChild(self.document.createTextNode(text))
            inline.appendChild(content)
            cell.appendChild(inline)

    @staticmethod
    def is_text_cell(cell):
        return cell is not None and cell.getAttribute("t") in ("s", "inlineStr")

    def update(self, number, changes):
        """changes: {열 번호: 새 값}. 바뀐 칸만 전달한다."""
        row = self.get_row(number)
        if row.hasAttribute("spans"):
            row.removeAttribute("spans")
        for column, text in changes.items():
            if self.locked_reason(number, column):
                continue
            cell = self.get_cell(row, column)
            self.set_value(cell, text, self.is_text_cell(cell))

    def insert(self, values, last_row, has_template):
        """마지막 데이터 행 아래에 추가하고 새 행 번호를 돌려준다."""
        number = last_row + 1
        while any(top <= number <= bottom for top, _, bottom, _ in self.merges):
            number += 1
        if number > 1048576:
            raise ExcelError("더 이상 행을 추가할 수 없습니다.")

        # 마지막 데이터 행의 서식과 자료형을 따른다.
        template = self.find_row(last_row) if has_template else None
        row = self.get_row(number)
        if template is not None:
            for name in ("s", "customFormat", "ht", "customHeight"):
                if template.hasAttribute(name):
                    row.setAttribute(name, template.getAttribute(name))
        if row.hasAttribute("spans"):
            row.removeAttribute("spans")

        for column, text in values.items():
            source = self.find_cell(template, column)
            cell = self.get_cell(row, column)
            if self.elements(cell, "f"):
                continue
            if source is not None and source.hasAttribute("s"):
                cell.setAttribute("s", source.getAttribute("s"))
            self.set_value(cell, text, self.is_text_cell(source))

        for dimension in self.elements(self.document, "dimension"):
            start, _, end = dimension.getAttribute("ref").partition(":")
            end = end or start
            columns = [column_number(start), column_number(end), *values]
            dimension.setAttribute(
                "ref",
                f"{column_name(min(columns))}{row_of(start)}:"
                f"{column_name(max(columns))}{max(row_of(end), number)}",
            )
        return number

    def delete(self, number):
        """행의 내용을 지운다. 수식 참조가 어긋나지 않도록 아래 행은 당기지 않는다."""
        row = self.find_row(number)
        if row is None:
            raise ExcelError("삭제할 행을 찾을 수 없습니다.")
        for formula in self.elements(row, "f"):
            ref = formula.getAttribute("ref")
            if ref and any(row_of(part) != number for part in ref.split(":")):
                raise ExcelError(
                    "다른 행과 함께 쓰는 수식이 있는 행이라 삭제할 수 없습니다.\n"
                    "Excel에서 직접 삭제해 주세요."
                )
        if self.elements(row, "f"):
            self.drop_calc_chain = True
        self.sheet_data.removeChild(row)

    def save(self):
        """임시 파일에 모두 쓴 뒤 원본과 바꾼다. 실패하면 원본은 그대로다."""
        handle, temp_path = tempfile.mkstemp(
            suffix=".tmp", dir=os.path.dirname(os.path.abspath(self.path))
        )
        os.close(handle)
        try:
            with zipfile.ZipFile(self.path) as source, \
                    zipfile.ZipFile(temp_path, "w") as target:
                for info in source.infolist():
                    name = info.filename
                    if self.drop_calc_chain and name == "xl/calcChain.xml":
                        continue
                    if name == self.sheet_path:
                        data = self.document.toxml(encoding="UTF-8")
                    else:
                        data = source.read(info)
                        if name == "xl/workbook.xml":
                            data = request_recalculation(data)
                        elif self.drop_calc_chain and name in (
                            "[Content_Types].xml", "xl/_rels/workbook.xml.rels"
                        ):
                            data = re.sub(
                                rb"<(?:\w+:)?(?:Override|Relationship)\b[^>]*calcChain\.xml[^>]*/>",
                                b"",
                                data,
                            )
                    target.writestr(info, data)
            os.replace(temp_path, self.path)
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)


def request_recalculation(workbook_xml):
    """Excel에서 열 때 수식을 다시 계산하도록 표시한다."""
    document = minidom.parseString(workbook_xml)
    properties = document.getElementsByTagNameNS(MAIN_NS, "calcPr")
    if not properties:
        return workbook_xml
    properties[0].setAttribute("fullCalcOnLoad", "1")
    return document.toxml(encoding="UTF-8")


# ---------------------------------------------------------------- 설정 파일

def read_config():
    """설정 내용. 파일이 없거나 잘못되었으면 빈 설정."""
    try:
        data = json.loads(CONFIG_FILE.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def write_config(data):
    try:
        CONFIG_FILE.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except OSError:
        pass  # 저장에 실패해도 파일 보기는 계속할 수 있어야 한다.


def load_last_path():
    """저장해 둔 파일 경로. 없으면 None."""
    path = read_config().get("last_file")
    return path if isinstance(path, str) and path else None


def save_last_path(path):
    data = read_config()  # 다른 설정 항목은 유지한다.
    data["last_file"] = os.path.abspath(path)
    write_config(data)


def file_signature(path):
    info = os.stat(path)
    return info.st_mtime_ns, info.st_size


# ---------------------------------------------------------------- 화면

class RecordForm(tk.Toplevel):
    """행 하나를 입력, 조회, 수정하는 폼."""

    def __init__(self, master, title, fields, on_save=None, on_edit=None):
        """fields: [(열 번호, 항목 이름, 값, 잠금 이유)]
        on_save가 없으면 조회 전용이다.
        """
        super().__init__(master)
        self.title(title)
        self.geometry("560x480")
        self.minsize(420, 240)
        self.transient(master)
        self.on_save = on_save
        self.inputs = {}  # 열 번호 -> (입력 위젯, 처음 값)

        buttons = ttk.Frame(self, padding=10)
        buttons.pack(side="bottom", fill="x")
        close_text = "취소" if on_save else "닫기"
        ttk.Button(buttons, text=close_text, command=self.destroy).pack(side="right")
        if on_save:
            ttk.Button(buttons, text="저장", command=self.save).pack(side="right", padx=6)
            self.bind("<Control-Return>", lambda _: self.save())
        elif on_edit:
            def edit():
                self.destroy()
                on_edit()
            ttk.Button(buttons, text="수정", command=edit).pack(side="right", padx=6)
        self.bind("<Escape>", lambda _: self.destroy())

        # 항목이 많아도 볼 수 있도록 스크롤 영역에 배치
        canvas = tk.Canvas(self, highlightthickness=0)
        scroll = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        body = ttk.Frame(canvas, padding=12)
        window = canvas.create_window(0, 0, window=body, anchor="nw")
        body.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window, width=e.width))

        def wheel(event):
            if body.winfo_height() > canvas.winfo_height():
                canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

        self.bind("<MouseWheel>", wheel)
        body.columnconfigure(1, weight=1)

        for index, (column, label, value, locked) in enumerate(fields):
            if locked and on_save:
                label = f"{label} ({locked})"
            ttk.Label(body, text=label, wraplength=160).grid(
                row=index, column=0, sticky="nw", padx=(0, 10), pady=4
            )
            readonly = on_save is None or locked is not None
            lines = value.count("\n") + 1
            widget = tk.Text(
                body, height=min(max(lines, 1), 5), width=30, wrap="word",
                font="TkDefaultFont", relief="solid", borderwidth=1, undo=True,
                padx=4, pady=3,
            )
            widget.insert("1.0", value)
            widget.edit_reset()
            if readonly:
                widget.configure(state="disabled", background="#eeeeee")
            else:
                # Tab 키로 다음 칸으로 이동
                widget.bind("<Tab>", self.focus_next)
                widget.bind("<Shift-Tab>", self.focus_previous)
                self.inputs[column] = (widget, value)
            widget.grid(row=index, column=1, sticky="ew", pady=4)

        self.grab_set()
        if self.inputs:
            next(iter(self.inputs.values()))[0].focus_set()
        else:
            self.focus_set()

    @staticmethod
    def focus_next(event):
        event.widget.tk_focusNext().focus_set()
        return "break"

    @staticmethod
    def focus_previous(event):
        event.widget.tk_focusPrev().focus_set()
        return "break"

    def save(self):
        values = {c: w.get("1.0", "end-1c") for c, (w, _) in self.inputs.items()}
        changed = {c: v for c, v in values.items() if v != self.inputs[c][1]}
        if self.on_save(self, values, changed):
            self.destroy()
        return "break"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("엑셀 첫 번째 시트 관리")
        self.geometry("900x600")

        self.path = None        # 현재 열린 파일
        self.signature = None   # 읽은 시점의 파일 상태
        self.rows = {}          # 행 번호 -> {열 번호: 값}
        self.columns = []
        self.header_row = None
        self.backed_up = set()  # 백업을 만든 파일

        style = ttk.Style(self)
        style.theme_use("clam")  # 선택 색상이 확실히 적용되는 테마
        style.configure("Treeview", rowheight=26)
        style.map(
            "Treeview",
            background=[("selected", "#1565c0")],
            foreground=[("selected", "#ffffff")],
        )

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Button(top, text="엑셀 파일 선택", command=self.choose_file).pack(side="left")
        self.status = ttk.Label(top, text="엑셀 파일을 선택하세요.")
        self.status.pack(side="left", padx=10)

        actions = ttk.Frame(self, padding=(10, 0, 10, 8))
        actions.pack(fill="x")
        self.buttons = {}
        for key, text, command in (
            ("new", "입력", self.create_record),
            ("view", "조회", self.view_record),
            ("edit", "수정", self.edit_record),
            ("delete", "삭제", self.delete_record),
            ("reload", "새로고침", self.reload),
        ):
            button = ttk.Button(actions, text=text, command=command)
            button.pack(side="left", padx=(0, 6))
            self.buttons[key] = button
        self.selection_label = ttk.Label(actions, foreground="#1565c0")
        self.selection_label.pack(side="left", padx=10)

        frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        frame.pack(fill="both", expand=True)
        self.table = ttk.Treeview(frame, show="headings", selectmode="browse")
        self.table.tag_configure("fixed", foreground="#777777")  # 제목 행과 그 위
        y_scroll = ttk.Scrollbar(frame, orient="vertical", command=self.table.yview)
        x_scroll = ttk.Scrollbar(frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.table.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self.table.bind("<<TreeviewSelect>>", lambda _: self.update_buttons())
        self.table.bind("<Double-1>", self.on_double_click)
        self.table.bind("<Return>", lambda _: self.view_record())
        self.table.bind("<Delete>", lambda _: self.delete_record())
        self.bind("<Control-n>", lambda _: self.create_record())

        self.update_buttons()
        self.after(100, self.load_last_file)  # 창을 먼저 띄운 뒤 불러온다.

    # ---- 파일 열기

    def load_last_file(self):
        if not CONFIG_FILE.exists():
            write_config({"last_file": ""})  # 설정 파일이 없으면 새로 만든다.
        path = load_last_path()
        if path is None:
            self.ask_to_choose("기억해 둔 엑셀 파일이 없습니다.")
        elif not os.path.isfile(path):
            self.ask_to_choose(
                "지난번에 연 엑셀 파일을 찾을 수 없습니다.\n"
                "경로가 바뀌었거나 파일이 삭제되었을 수 있습니다.\n\n"
                f"{path}"
            )
        else:
            self.load(path)

    def ask_to_choose(self, reason):
        self.status.configure(text="엑셀 파일을 선택해 주세요.")
        messagebox.showinfo(
            "파일 선택 필요",
            f"{reason}\n\n'엑셀 파일 선택' 버튼을 눌러 파일을 선택해 주세요.",
            parent=self,
        )

    def choose_file(self):
        last = load_last_path()
        folder = os.path.dirname(last) if last else ""
        path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            initialdir=folder if os.path.isdir(folder) else None,
            filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")],
        )
        if path:
            save_last_path(path)  # 선택하는 즉시 전체 경로를 저장한다.
            self.load(path)

    def load(self, path, select=None):
        self.clear()
        try:
            signature = file_signature(path)
            name, rows = read_first_sheet(path)
        except ExcelError as error:
            self.show_error(str(error))
        except Exception as error:  # 손상된 파일, 접근 권한 문제 등
            self.show_error(f"파일을 읽는 중 오류가 발생했습니다.\n{error}")
        else:
            self.path, self.signature = path, signature
            self.show(os.path.basename(path), name, rows)
            if select is not None and self.table.exists(str(select)):
                self.table.selection_set(str(select))
                self.table.focus(str(select))
                self.table.see(str(select))
                self.table.focus_set()
        self.update_buttons()

    def reload(self):
        if self.path:
            self.load(self.path, self.selected_row())

    def clear(self):
        self.path = self.signature = self.header_row = None
        self.rows, self.columns = {}, []
        self.table.delete(*self.table.get_children())
        self.table.configure(columns=())
        self.status.configure(text="엑셀 파일을 선택하세요.")
        self.update_buttons()

    def show_error(self, message):
        self.status.configure(text="파일을 열지 못했습니다.")
        messagebox.showerror("열기 실패", message)

    def show(self, file_name, name, rows):
        if not rows:
            self.status.configure(text=f'{file_name} · 시트 "{name}"에 표시할 내용이 없습니다.')
            return

        self.rows = dict(rows)
        # 값이 2개 이상인 첫 행을 제목 행으로 본다. (한 칸짜리 문서 제목은 건너뜀)
        self.header_row = next((n for n, cells in rows if len(cells) >= 2), rows[0][0])
        self.columns = sorted(
            {c for n, cells in rows if n >= self.header_row for c in cells}
        )
        shown = sorted({c for _, cells in rows for c in cells})
        ids = ["row"] + [column_name(c) for c in shown]
        self.table.configure(columns=ids)
        self.table.heading("row", text="행")
        self.table.column("row", width=50, stretch=False)
        for column in shown:
            column_id = column_name(column)
            title = self.rows[self.header_row].get(column, column_id)
            self.table.heading(column_id, text=" ".join(title.splitlines()))
            self.table.column(column_id, width=130, stretch=False)

        # 제목 행은 열 이름으로 보여주므로 목록에서는 제외한다.
        visible = [(n, cells) for n, cells in rows if n != self.header_row]
        for row_number, cells in visible[:MAX_ROWS]:
            # 목록은 한 줄로 표시하므로 셀 안의 줄바꿈은 공백으로 바꾼다.
            values = [" ".join(cells.get(c, "").splitlines()) for c in shown]
            tags = ("fixed",) if row_number < self.header_row else ()
            self.table.insert(
                "", "end", iid=str(row_number), values=[row_number] + values, tags=tags
            )

        count = sum(1 for n in self.rows if n > self.header_row)
        limit = f" (처음 {MAX_ROWS}행만 표시)" if len(visible) > MAX_ROWS else ""
        self.status.configure(
            text=f'{file_name} · 시트 "{name}" · 제목 {self.header_row}행 · 데이터 {count}행{limit}'
        )

    # ---- 선택

    def selected_row(self):
        selection = self.table.selection()
        return int(selection[0]) if selection else None

    def update_buttons(self):
        number = self.selected_row()
        editable = number is not None and number > self.header_row
        states = {
            "new": self.header_row is not None,
            "view": number is not None,
            "edit": editable,
            "delete": editable,
            "reload": self.path is not None,
        }
        for key, enabled in states.items():
            self.buttons[key].state(["!disabled" if enabled else "disabled"])
        if number is None:
            text = "선택된 행이 없습니다."
        else:
            note = "" if editable else " (제목 영역은 수정할 수 없음)"
            text = f"▶ {number}행 선택됨{note}"
        self.selection_label.configure(text=text)

    def on_double_click(self, event):
        if self.table.identify_row(event.y):
            self.view_record()

    # ---- CRUD

    def file_unchanged(self):
        """다른 프로그램이 파일을 바꿨으면 다시 읽고 False를 돌려준다."""
        try:
            if file_signature(self.path) == self.signature:
                return True
        except OSError:
            pass
        messagebox.showwarning(
            "파일 변경됨",
            "파일이 다른 곳에서 변경되었거나 사라졌습니다.\n"
            "다시 불러오니 확인한 뒤 작업해 주세요.",
        )
        self.reload()
        return False

    def open_editor(self):
        try:
            return SheetEditor(self.path)
        except ExcelError as error:
            messagebox.showerror("작업 실패", str(error))
        except Exception as error:
            messagebox.showerror("작업 실패", f"파일을 열 수 없습니다.\n{error}")
        return None

    def fields(self, editor, number, source):
        """폼에 보여줄 항목 목록. number는 잠금 확인용(새 행이면 None)."""
        header = self.rows[self.header_row]
        result = []
        for column in self.columns:
            name = header.get(column)
            label = " ".join(name.splitlines()) if name else f"{column_name(column)}열"
            locked = None if number is None else editor.locked_reason(number, column)
            result.append((column, label, source.get(column, ""), locked))
        return result

    def commit(self, parent, action):
        """action(editor)로 파일을 수정해 저장한다. (성공 여부, action의 결과)"""
        try:
            if file_signature(self.path) != self.signature:
                raise ExcelError(
                    "파일이 다른 곳에서 변경되었습니다.\n"
                    "입력한 내용을 복사해 두고, 폼을 닫은 뒤 새로고침해 주세요."
                )
            editor = SheetEditor(self.path)
            result = action(editor)
            if self.path not in self.backed_up:
                original = Path(self.path)
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                shutil.copy2(
                    original,
                    original.with_name(f"{original.stem}_백업_{stamp}{original.suffix}"),
                )
                self.backed_up.add(self.path)
            editor.save()
        except ExcelError as error:
            messagebox.showerror("저장 실패", str(error), parent=parent)
        except PermissionError:
            messagebox.showerror(
                "저장 실패",
                "파일에 저장할 수 없습니다.\n"
                "Excel에서 열려 있다면 닫고 다시 시도해 주세요.",
                parent=parent,
            )
        except Exception as error:
            messagebox.showerror("저장 실패", f"저장 중 오류가 발생했습니다.\n{error}", parent=parent)
        else:
            return True, result
        return False, None

    def create_record(self):
        if self.header_row is None or not self.file_unchanged():
            return

        def save(form, values, _):
            filled = {c: v for c, v in values.items() if v.strip()}
            if not filled:
                messagebox.showwarning("입력 필요", "값을 하나 이상 입력해 주세요.", parent=form)
                return False
            last = max(self.rows)
            ok, number = self.commit(
                form, lambda e: e.insert(filled, last, last > self.header_row)
            )
            if ok:
                self.load(self.path, number)
            return ok

        RecordForm(self, "새 항목 입력", self.fields(None, None, {}), on_save=save)

    def view_record(self):
        number = self.selected_row()
        if number is None or not self.file_unchanged():
            return
        on_edit = self.edit_record if number > self.header_row else None
        RecordForm(
            self,
            f"개별 조회 - {number}행",
            [(c, label, value, None) for c, label, value, _
             in self.fields(None, None, self.rows[number])],
            on_edit=on_edit,
        )

    def edit_record(self):
        number = self.selected_row()
        if number is None or number <= self.header_row or not self.file_unchanged():
            return
        editor = self.open_editor()
        if editor is None:
            return
        original = self.rows[number]

        def save(form, values, changed):
            if not changed:
                return True
            if not any(v.strip() for v in {**original, **values}.values()):
                messagebox.showwarning(
                    "저장할 수 없음",
                    "모든 값을 비울 수 없습니다. 행을 없애려면 삭제를 사용해 주세요.",
                    parent=form,
                )
                return False
            ok, _ = self.commit(form, lambda e: e.update(number, changed))
            if ok:
                self.load(self.path, number)
            return ok

        RecordForm(
            self, f"수정 - {number}행", self.fields(editor, number, original), on_save=save
        )

    def delete_record(self):
        number = self.selected_row()
        if number is None or number <= self.header_row or not self.file_unchanged():
            return
        preview = "\n".join(
            f"· {label}: {' '.join(value.split())[:40]}"
            for _, label, value, _ in self.fields(None, None, self.rows[number])[:8]
            if value
        )
        confirmed = messagebox.askyesno(
            "삭제 확인",
            f"{number}행을 정말 삭제하시겠습니까?\n\n{preview}",
            icon="warning",
            default="no",
            parent=self,
        )
        if not confirmed:
            return
        ok, _ = self.commit(self, lambda e: e.delete(number))
        if ok:
            self.load(self.path)


if __name__ == "__main__":
    App().mainloop()
