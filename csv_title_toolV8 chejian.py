# -*- coding: utf-8 -*-

import csv
import os
import re
import zipfile
import posixpath
import tempfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QLineEdit, QTextEdit, QListWidget, QListWidgetItem,
    QFileDialog, QMessageBox, QVBoxLayout, QHBoxLayout,
    QGridLayout, QGroupBox, QCheckBox, QRadioButton,
    QButtonGroup, QProgressBar, QFrame, QScrollArea,
    QSizePolicy
)


# ============================================================
# 1. CSV 读取
# ============================================================

def detect_csv_format(filename):
    encodings = ["utf-8-sig", "utf-8", "gb18030", "big5"]

    encoding = None

    with open(filename, "rb") as f:
        sample_bytes = f.read(65536)

    for item in encodings:
        try:
            sample_bytes.decode(item)
            encoding = item
            break
        except UnicodeDecodeError:
            pass

    if encoding is None:
        raise ValueError(f"无法识别文件编码：{filename}")

    sample_text = sample_bytes.decode(
        encoding,
        errors="replace"
    )

    try:
        dialect = csv.Sniffer().sniff(
            sample_text[:8192],
            delimiters=",;\t|"
        )
    except csv.Error:
        dialect = csv.excel

    return encoding, dialect


def open_csv_reader(filename):
    encoding, dialect = detect_csv_format(filename)

    file_obj = open(
        filename,
        "r",
        encoding=encoding,
        newline=""
    )

    return file_obj, csv.reader(file_obj, dialect)


# ============================================================
# 2. XLSX 读取
# ============================================================

def local_name(tag):
    return tag.split("}")[-1]


def get_xlsx_sheet_info(archive):
    names = archive.namelist()

    if "xl/workbook.xml" not in names:
        raise ValueError("不是有效的 XLSX 文件。")

    workbook = ET.fromstring(
        archive.read("xl/workbook.xml")
    )

    first_sheet = next(
        (
            item for item in workbook.iter()
            if local_name(item.tag) == "sheet"
        ),
        None
    )

    if first_sheet is None:
        raise ValueError("Excel 文件中没有工作表。")

    relation_id = next(
        (
            value for key, value in first_sheet.attrib.items()
            if local_name(key) == "id"
        ),
        None
    )

    if not relation_id:
        raise ValueError("无法识别 Excel 工作表。")

    rels_path = "xl/_rels/workbook.xml.rels"

    relationships = ET.fromstring(
        archive.read(rels_path)
    )

    target = next(
        (
            rel.attrib.get("Target")
            for rel in relationships
            if rel.attrib.get("Id") == relation_id
        ),
        None
    )

    if not target:
        raise ValueError("无法定位 Excel 工作表。")

    if target.startswith("/"):
        sheet_path = target.lstrip("/")
    else:
        sheet_path = posixpath.normpath(
            posixpath.join("xl", target)
        )

    if sheet_path not in names:
        raise ValueError("找不到 Excel 工作表：" + sheet_path)

    shared_strings = []

    if "xl/sharedStrings.xml" in names:
        root = ET.fromstring(
            archive.read("xl/sharedStrings.xml")
        )

        for item in root:
            if local_name(item.tag) == "si":
                shared_strings.append(
                    "".join(
                        node.text or ""
                        for node in item.iter()
                        if local_name(node.tag) == "t"
                    )
                )

    return sheet_path, shared_strings


def excel_col_index(cell_ref):
    match = re.match(r"([A-Z]+)", cell_ref or "")

    if not match:
        return 0

    number = 0

    for char in match.group(1):
        number = number * 26 + ord(char) - 64

    return number - 1


def xlsx_cell_value(cell, shared_strings):
    cell_type = cell.attrib.get("t", "")
    value = ""

    for child in cell:
        kind = local_name(child.tag)

        if kind == "v":
            value = child.text or ""
            break

        if kind == "is":
            value = "".join(
                node.text or ""
                for node in child.iter()
                if local_name(node.tag) == "t"
            )
            break

    if cell_type == "s" and value:
        try:
            value = shared_strings[int(value)]
        except (ValueError, IndexError):
            value = ""

    elif cell_type == "b":
        value = "TRUE" if value == "1" else "FALSE"

    return value


def iter_xlsx_rows(filename):
    try:
        with zipfile.ZipFile(filename) as archive:

            sheet_path, shared_strings = get_xlsx_sheet_info(
                archive
            )

            with archive.open(sheet_path) as sheet_file:

                stack = []

                for event, element in ET.iterparse(
                    sheet_file,
                    events=("start", "end")
                ):

                    if event == "start":
                        stack.append(element)
                        continue

                    if local_name(element.tag) == "row":

                        cells = {}
                        max_index = -1

                        for cell in element:

                            if local_name(cell.tag) != "c":
                                continue

                            index = excel_col_index(
                                cell.attrib.get("r", "")
                            )

                            cells[index] = xlsx_cell_value(
                                cell,
                                shared_strings
                            )

                            max_index = max(max_index, index)

                        row = [""] * (max_index + 1)

                        for index, value in cells.items():
                            row[index] = value

                        yield row

                        element.clear()

                        if len(stack) >= 2:
                            stack[-2].remove(element)

                    stack.pop()

    except zipfile.BadZipFile:
        raise ValueError(
            f"{filename} 不是有效的 XLSX 文件。"
        )


def read_xlsx_first_column(filename):
    values = []

    for row in iter_xlsx_rows(filename):
        values.append(row[0] if row else "")

    return values


# ============================================================
# 3. XLSX 输出
# ============================================================

class XlsxStreamingWriter:

    def __init__(self, filename):
        self.filename = filename

        self.temp_dir = tempfile.TemporaryDirectory(
            prefix="title_xlsx_"
        )

        self.sheet_path = os.path.join(
            self.temp_dir.name,
            "sheet1.xml"
        )

        self.sheet = open(
            self.sheet_path,
            "w",
            encoding="utf-8",
            newline=""
        )

        self.row_number = 0

        self.sheet.write(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        )

        self.sheet.write(
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        )

    def writerow(self, values):
        self.row_number += 1

        if self.row_number > 1048576:
            raise ValueError(
                "XLSX 超过最大行数 1,048,576 行，请改用 CSV 输出。"
            )

        self.sheet.write(
            f'<row r="{self.row_number}">'
        )

        for col_number, value in enumerate(values, 1):

            if col_number > 16384:
                raise ValueError(
                    "XLSX 超过最大列数 16,384 列。"
                )

            n = col_number
            letters = ""

            while n:
                n, remainder = divmod(n - 1, 26)
                letters = chr(65 + remainder) + letters

            text_value = "" if value is None else str(value)

            preserve = (
                ' xml:space="preserve"'
                if text_value != text_value.strip()
                else ""
            )

            self.sheet.write(
                f'<c r="{letters}{self.row_number}" t="inlineStr">'
                f'<is><t{preserve}>{escape(text_value)}</t></is></c>'
            )

        self.sheet.write("</row>")

    def close(self):
        if self.sheet.closed:
            return

        self.sheet.write("</sheetData></worksheet>")
        self.sheet.close()

        with zipfile.ZipFile(
            self.filename,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=1
        ) as archive:

            archive.writestr(
                "[Content_Types].xml",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
            )

            archive.writestr(
                "_rels/.rels",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
            )

            archive.writestr(
                "xl/workbook.xml",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="提取结果" sheetId="1" r:id="rId1"/></sheets>
</workbook>"""
            )

            archive.writestr(
                "xl/_rels/workbook.xml.rels",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
            )

            archive.write(
                self.sheet_path,
                "xl/worksheets/sheet1.xml"
            )

        self.temp_dir.cleanup()

    def abort(self):
        try:
            self.sheet.close()
        finally:
            self.temp_dir.cleanup()


# ============================================================
# 4. 读取查询清单
# ============================================================

def clean_queries(values):
    queries = []

    for value in values:
        if value is None:
            continue

        title = str(value).strip()

        if title:
            queries.append(title)

    return list(dict.fromkeys(queries))


def read_query_file(filename, skip_header=False):
    extension = os.path.splitext(filename)[1].lower()

    if extension == ".csv":

        file_obj, reader = open_csv_reader(filename)

        try:
            values = [
                row[0] if row else ""
                for row in reader
            ]
        finally:
            file_obj.close()

    elif extension == ".xlsx":
        values = read_xlsx_first_column(filename)

    else:
        raise ValueError("查询清单只支持 CSV、XLSX。")

    if skip_header and values:
        values = values[1:]

    return clean_queries(values)


# ============================================================
# 5. 核心提取程序
# ============================================================

def process_files_streaming(
    filenames,
    queries,
    output_path,
    progress_callback=None
):
    found = set()
    errors = []

    matched_files = 0
    matched_columns = 0
    output_rows = 0

    is_xlsx_output = output_path.lower().endswith(".xlsx")

    writer_obj = None
    output_file = None

    try:

        if is_xlsx_output:
            writer_obj = XlsxStreamingWriter(output_path)
            writer = writer_obj

        else:
            output_file = open(
                output_path,
                "w",
                encoding="utf-8-sig",
                newline=""
            )

            writer = csv.writer(output_file)

        writer.writerow(queries)

        for file_number, filename in enumerate(filenames, 1):

            display_name = os.path.basename(filename)

            message = (
                f"正在处理第 {file_number}/{len(filenames)} 个文件："
                f"{display_name}"
            )

            if progress_callback:
                progress_callback(
                    file_number,
                    len(filenames),
                    message
                )

            file_obj = None

            try:

                if filename.lower().endswith(".csv"):
                    file_obj, reader = open_csv_reader(filename)
                    row_iterator = iter(reader)

                elif filename.lower().endswith(".xlsx"):
                    row_iterator = iter_xlsx_rows(filename)

                else:
                    raise ValueError("不支持的文件格式。")

                try:
                    titles = next(row_iterator)
                except StopIteration:
                    raise ValueError("文件为空。")

                first_index_by_title = {}

                for index, title in enumerate(titles):
                    if title not in first_index_by_title:
                        first_index_by_title[title] = index

                title_indexes = {}

                for query in queries:

                    if query in first_index_by_title:
                        title_indexes[query] = first_index_by_title[query]
                        found.add(query)
                        matched_columns += 1

                if title_indexes:
                    matched_files += 1

                for row in row_iterator:

                    output_row = []

                    for query in queries:

                        index = title_indexes.get(query)

                        if index is not None and index < len(row):
                            output_row.append(row[index])
                        else:
                            output_row.append("")

                    writer.writerow(output_row)
                    output_rows += 1

            except Exception as error:

                errors.append(
                    f"{display_name}：{error}"
                )

            finally:

                if file_obj:
                    file_obj.close()

        if writer_obj:
            writer_obj.close()
            writer_obj = None

    except Exception:

        if writer_obj:
            writer_obj.abort()

        raise

    finally:

        if output_file:
            output_file.close()

    missing = [
        query for query in queries
        if query not in found
    ]

    return {
        "query_count": len(queries),
        "matched_files": matched_files,
        "matched_columns": matched_columns,
        "output_rows": output_rows,
        "missing": missing,
        "errors": errors
    }


# ============================================================
# 6. 后台处理线程
# ============================================================

class ExtractWorker(QThread):

    progress = Signal(int, int, str)
    succeeded = Signal(dict)
    failed = Signal(str)

    def __init__(self, filenames, queries, output_path):
        super().__init__()

        self.filenames = list(filenames)
        self.queries = list(queries)
        self.output_path = output_path

    def run(self):
        try:

            result = process_files_streaming(
                self.filenames,
                self.queries,
                self.output_path,
                progress_callback=lambda current, total, message:
                    self.progress.emit(current, total, message)
            )

            result["output_path"] = self.output_path

            self.succeeded.emit(result)

        except Exception as exc:
            self.failed.emit(str(exc))


# ============================================================
# 7. 主界面
# ============================================================

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("TITLE 超大型文件提取工作台")

        self.resize(1250, 900)
        self.setMinimumSize(850, 650)

        self.data_files = []
        self.query_file = ""
        self.worker = None
        self.query_mode = "file"

        self._build_ui()
        self._apply_style()

        self._log("程序已启动，等待选择文件。")

    # --------------------------------------------------------
    # 创建界面
    # --------------------------------------------------------

    def _build_ui(self):

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )

        scroll.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )

        central = QWidget()

        central.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Preferred
        )

        scroll.setWidget(central)
        self.setCentralWidget(scroll)

        root = QVBoxLayout(central)

        root.setContentsMargins(22, 18, 22, 18)
        root.setSpacing(13)

        # 顶部标题

        header = QFrame()
        header.setObjectName("header")

        header_layout = QVBoxLayout(header)

        title = QLabel("TITLE 超大型文件提取工作台")
        title.setObjectName("mainTitle")

        subtitle = QLabel(
            "MAC EDITION  ·  原生桌面界面  ·  CSV / XLSX  ·  大文件逐行处理"
        )

        subtitle.setObjectName("subtitle")

        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)

        root.addWidget(header)

        # 01 原始文件

        data_group = QGroupBox("01  选择原始数据文件")

        data_layout = QVBoxLayout(data_group)

        hint = QLabel(
            "支持 CSV、XLSX，可混合选择多个文件；支持分批添加、单独删除。"
        )

        hint.setObjectName("hint")
        data_layout.addWidget(hint)

        buttons = QHBoxLayout()

        self.add_data_btn = QPushButton(
            "📁 添加原始文件（CSV / XLSX）"
        )
        self.add_data_btn.setObjectName("blueButton")

        self.remove_data_btn = QPushButton("删除选中文件")
        self.remove_data_btn.setObjectName("redButton")

        self.clear_data_btn = QPushButton("清空全部文件")
        self.clear_data_btn.setObjectName("grayButton")

        buttons.addWidget(self.add_data_btn)
        buttons.addWidget(self.remove_data_btn)
        buttons.addWidget(self.clear_data_btn)
        buttons.addStretch()

        data_layout.addLayout(buttons)

        self.data_list = QListWidget()

        self.data_list.setSelectionMode(
            QListWidget.SelectionMode.ExtendedSelection
        )

        self.data_list.setMinimumHeight(145)

        self.data_list.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding
        )

        data_layout.addWidget(self.data_list, 1)

        self.data_count = QLabel("当前已添加 0 个文件")
        self.data_count.setObjectName("summary")

        data_layout.addWidget(self.data_count)

        data_group.setMinimumHeight(205)

        root.addWidget(data_group)

        # 02 TITLE 查询内容

        query_group = QGroupBox("02  设置 TITLE 查询内容")

        query_layout = QVBoxLayout(query_group)

        hint = QLabel(
            "查询方式二选一：从 CSV / XLSX 文件导入，或直接手动输入 TITLE。"
        )

        hint.setObjectName("hint")
        query_layout.addWidget(hint)

        mode_row = QHBoxLayout()

        self.file_mode_radio = QRadioButton("📁 文件导入")
        self.manual_mode_radio = QRadioButton("✍️ 手动输入 TITLE")

        self.file_mode_radio.setChecked(True)

        self.mode_group = QButtonGroup(self)

        self.mode_group.addButton(self.file_mode_radio)
        self.mode_group.addButton(self.manual_mode_radio)

        mode_row.addWidget(self.file_mode_radio)
        mode_row.addWidget(self.manual_mode_radio)
        mode_row.addStretch()

        query_layout.addLayout(mode_row)

        # 文件导入面板

        self.file_query_panel = QWidget()

        file_layout = QVBoxLayout(self.file_query_panel)
        file_layout.setContentsMargins(0, 4, 0, 0)

        self.choose_query_btn = QPushButton(
            "📊 选择 TITLE 查询文件"
        )

        self.choose_query_btn.setObjectName("purpleButton")

        file_layout.addWidget(
            self.choose_query_btn,
            alignment=Qt.AlignmentFlag.AlignLeft
        )

        self.query_path_edit = QLineEdit()
        self.query_path_edit.setReadOnly(True)

        self.query_path_edit.setPlaceholderText(
            "尚未选择查询清单"
        )

        file_layout.addWidget(self.query_path_edit)

        self.skip_header = QCheckBox(
            "忽略查询清单第一行（第一行是字段名称时勾选）"
        )

        file_layout.addWidget(self.skip_header)

        query_layout.addWidget(self.file_query_panel)

        # 手动输入面板

        self.manual_query_panel = QWidget()

        manual_layout = QVBoxLayout(self.manual_query_panel)
        manual_layout.setContentsMargins(0, 4, 0, 0)

        self.manual_titles = QTextEdit()

        self.manual_titles.setPlaceholderText(
            "每行输入一个 TITLE，例如：\n"
            "Device_SN\n"
            "Product_Name\n"
            "Test_Result\n"
            "Serial_Number"
        )

        self.manual_titles.setMinimumHeight(125)

        self.manual_titles.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding
        )

        manual_layout.addWidget(self.manual_titles)

        manual_bottom = QHBoxLayout()

        self.manual_count = QLabel(
            "当前输入 0 个有效 TITLE"
        )

        self.manual_count.setObjectName("summary")

        self.clear_manual_btn = QPushButton("清空手动输入")
        self.clear_manual_btn.setObjectName("orangeButton")

        manual_bottom.addWidget(self.manual_count)
        manual_bottom.addStretch()
        manual_bottom.addWidget(self.clear_manual_btn)

        manual_layout.addLayout(manual_bottom)

        self.manual_query_panel.hide()

        query_layout.addWidget(self.manual_query_panel)

        rule = QLabel(
            "查询规则：读取 CSV 第一列或 XLSX 第一个工作表第一列；"
            "手动输入时每行一个 TITLE；"
            "自动去除空行、首尾空格和重复项；"
            "与原始文件第一行精确匹配；多个源文件的结果纵向合并。"
        )

        rule.setWordWrap(True)
        rule.setObjectName("rule")

        query_layout.addWidget(rule)

        query_group.setMinimumHeight(245)

        root.addWidget(query_group)

        # 03 输出设置

        output_group = QGroupBox("03  输出文件设置")

        output_layout = QGridLayout(output_group)

        output_layout.addWidget(
            QLabel("输出文件名称："),
            0, 0
        )

        self.output_name_edit = QLineEdit(
            "TITLE整列提取结果.csv"
        )

        output_layout.addWidget(
            self.output_name_edit,
            0, 1, 1, 2
        )

        output_layout.addWidget(
            QLabel("输出文件路径："),
            1, 0
        )

        self.output_path_edit = QLineEdit()

        self.output_path_edit.setPlaceholderText(
            "请选择输出文件保存位置"
        )

        output_layout.addWidget(
            self.output_path_edit,
            1, 1
        )

        self.browse_output_btn = QPushButton(
            "选择保存位置"
        )

        self.browse_output_btn.setObjectName("orangeButton")

        output_layout.addWidget(
            self.browse_output_btn,
            1, 2
        )

        note = QLabel(
            "默认 CSV；文件名以 .xlsx 结尾则导出 Excel；"
            "不填写扩展名时默认补充 .csv。"
        )

        note.setObjectName("hint")

        output_layout.addWidget(
            note,
            2, 0, 1, 3
        )

        output_group.setMinimumHeight(135)

        root.addWidget(output_group)

        # 04 批量提取

        run_group = QGroupBox("04  开始批量提取")

        run_layout = QVBoxLayout(run_group)

        self.run_btn = QPushButton("▶  开始批量提取")
        self.run_btn.setObjectName("greenButton")
        self.run_btn.setMinimumHeight(48)

        run_layout.addWidget(self.run_btn)

        self.status_label = QLabel("运行状态：等待操作")
        self.status_label.setObjectName("summary")

        run_layout.addWidget(self.status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)

        run_layout.addWidget(self.progress_bar)

        log_title = QLabel("处理日志")
        log_title.setObjectName("logTitle")

        run_layout.addWidget(log_title)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMinimumHeight(150)

        self.log_box.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding
        )

        self.log_box.setObjectName("logBox")

        run_layout.addWidget(self.log_box, 1)

        run_group.setMinimumHeight(330)

        run_group.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding
        )

        root.addWidget(run_group)

        root.addStretch(1)

        # 事件绑定

        self.add_data_btn.clicked.connect(
            self.add_data_files
        )

        self.remove_data_btn.clicked.connect(
            self.remove_selected_files
        )

        self.clear_data_btn.clicked.connect(
            self.clear_data_files
        )

        self.choose_query_btn.clicked.connect(
            self.choose_query_file
        )

        self.file_mode_radio.toggled.connect(
            self.switch_query_mode
        )

        self.manual_mode_radio.toggled.connect(
            self.switch_query_mode
        )

        self.manual_titles.textChanged.connect(
            self.update_manual_count
        )

        self.clear_manual_btn.clicked.connect(
            self.clear_manual_titles
        )

        self.browse_output_btn.clicked.connect(
            self.choose_output_path
        )

        self.run_btn.clicked.connect(
            self.start_extract
        )

    # --------------------------------------------------------
    # 界面样式
    # --------------------------------------------------------

    def _apply_style(self):

        self.setStyleSheet("""
            QMainWindow, QWidget {
                background:#f3f6fb;
                color:#24324a;
                font-size:13px;
            }

            QScrollArea {
                border:none;
                background:#f3f6fb;
            }

            QScrollArea > QWidget > QWidget {
                background:#f3f6fb;
            }

            #header {
                background:#203b63;
                border-radius:12px;
                padding:10px;
            }

            #mainTitle {
                color:white;
                font-size:23px;
                font-weight:800;
            }

            #subtitle {
                color:#d5e2f2;
                font-size:12px;
            }

            QGroupBox {
                background:white;
                border:1px solid #dfe6ef;
                border-radius:11px;
                margin-top:12px;
                padding:12px;
                font-weight:700;
            }

            QGroupBox::title {
                subcontrol-origin:margin;
                left:12px;
                padding:0 6px;
                color:#203b63;
            }

            QLabel {
                background:transparent;
            }

            #hint {
                color:#718096;
                font-weight:400;
            }

            #summary {
                color:#365579;
                font-weight:700;
            }

            #rule {
                background:#f2f7ff;
                border-left:4px solid #3478f6;
                border-radius:5px;
                padding:9px;
                color:#52647d;
            }

            QLineEdit, QTextEdit, QListWidget {
                background:#fbfcff;
                border:1px solid #d8e1ed;
                border-radius:7px;
                padding:7px;
                selection-background-color:#3478f6;
            }

            QLineEdit:focus,
            QTextEdit:focus,
            QListWidget:focus {
                border:1px solid #805ad5;
            }

            QPushButton {
                border:none;
                border-radius:7px;
                padding:10px 14px;
                color:white;
                font-weight:700;
            }

            QPushButton:disabled {
                background:#aab4c2;
                color:#eef1f5;
            }

            #blueButton {
                background:#3478f6;
            }

            #purpleButton {
                background:#805ad5;
            }

            #greenButton {
                background:#159b65;
                font-size:16px;
            }

            #orangeButton {
                background:#d98a2b;
            }

            #redButton {
                background:#d9534f;
            }

            #grayButton {
                background:#64748b;
            }

            #logTitle {
                font-size:14px;
                font-weight:800;
            }

            #logBox {
                background:#17263d;
                color:#cde0f6;
                font-family:Menlo,monospace;
            }

            QProgressBar {
                border:none;
                border-radius:5px;
                background:#e8edf4;
                height:12px;
                text-align:center;
            }

            QProgressBar::chunk {
                background:#19a66b;
                border-radius:5px;
            }

            QCheckBox, QRadioButton {
                spacing:7px;
            }
        """)

    # --------------------------------------------------------
    # 日志
    # --------------------------------------------------------

    def _log(self, message):
        self.log_box.append(str(message))

    # --------------------------------------------------------
    # 检查运行状态
    # --------------------------------------------------------

    def _ensure_idle(self):

        if self.worker and self.worker.isRunning():

            QMessageBox.information(
                self,
                "正在处理",
                "当前正在提取数据，请等待完成后再修改设置。"
            )

            return False

        return True

    # --------------------------------------------------------
    # 添加原始文件
    # --------------------------------------------------------

    def add_data_files(self):

        if not self._ensure_idle():
            return

        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "选择原始数据文件",
            "",
            "支持的文件 (*.csv *.xlsx);;CSV 文件 (*.csv);;Excel 文件 (*.xlsx)"
        )

        if not paths:
            return

        added = 0
        duplicate = 0

        for path in paths:

            if path in self.data_files:
                duplicate += 1
            else:
                self.data_files.append(path)
                added += 1

        self.refresh_data_list()

        self._log(
            f"本次添加 {added} 个文件，重复跳过 {duplicate} 个。"
        )

    # --------------------------------------------------------
    # 刷新原始文件列表
    # --------------------------------------------------------

    def refresh_data_list(self):

        self.data_list.clear()

        total_size = 0

        for index, path in enumerate(self.data_files, 1):

            try:
                size = os.path.getsize(path)
                total_size += size
                size_text = self.format_size(size)

            except OSError:
                size_text = "文件不存在"

            item = QListWidgetItem(
                f"{index}. {os.path.basename(path)}\n"
                f"{path}\n"
                f"文件大小：{size_text}"
            )

            item.setToolTip(path)

            self.data_list.addItem(item)

        self.data_count.setText(
            f"当前已添加 {len(self.data_files)} 个文件，"
            f"总大小：{self.format_size(total_size)}"
        )

    # --------------------------------------------------------
    # 文件大小转换
    # --------------------------------------------------------

    @staticmethod
    def format_size(size):

        if size < 1024:
            return f"{size} B"

        if size < 1024 ** 2:
            return f"{size / 1024:.1f} KB"

        if size < 1024 ** 3:
            return f"{size / 1024 ** 2:.2f} MB"

        return f"{size / 1024 ** 3:.2f} GB"

    # --------------------------------------------------------
    # 删除选中文件
    # --------------------------------------------------------

    def remove_selected_files(self):

        if not self._ensure_idle():
            return

        indexes = sorted(
            {
                self.data_list.row(item)
                for item in self.data_list.selectedItems()
            },
            reverse=True
        )

        if not indexes:

            QMessageBox.information(
                self,
                "提示",
                "请先选中要移除的文件。"
            )

            return

        names = [
            os.path.basename(self.data_files[i])
            for i in indexes
        ]

        answer = QMessageBox.question(
            self,
            "确认移除",
            "确定从待处理列表移除以下文件吗？\n\n"
            + "\n".join(names)
            + "\n\n只会从程序列表移除，不会删除硬盘上的原文件。",
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if answer != QMessageBox.StandardButton.Yes:
            return

        for index in indexes:
            self.data_files.pop(index)

        self.refresh_data_list()

        self._log(
            "已从待处理列表移除：" + "、".join(names)
        )

    # --------------------------------------------------------
    # 清空原始文件
    # --------------------------------------------------------

    def clear_data_files(self):

        if not self._ensure_idle():
            return

        if not self.data_files:

            QMessageBox.information(
                self,
                "提示",
                "当前列表已经是空的。"
            )

            return

        answer = QMessageBox.question(
            self,
            "确认清空",
            f"确定清空全部 {len(self.data_files)} 个待处理文件吗？\n"
            "不会删除硬盘上的原始文件。",
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if answer == QMessageBox.StandardButton.Yes:

            self.data_files.clear()
            self.refresh_data_list()

            self._log(
                "全部原始文件已从待处理列表移除。"
            )

    # --------------------------------------------------------
    # 选择查询清单
    # --------------------------------------------------------

    def choose_query_file(self):

        if not self._ensure_idle():
            return

        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 TITLE 查询清单",
            "",
            "支持的文件 (*.csv *.xlsx);;CSV 文件 (*.csv);;Excel 文件 (*.xlsx)"
        )

        if path:

            self.query_file = path
            self.query_path_edit.setText(path)

            self._log("查询清单：" + path)

    # --------------------------------------------------------
    # 切换查询模式
    # --------------------------------------------------------

    def switch_query_mode(self, checked):

        if not checked:
            return

        if self.file_mode_radio.isChecked():

            self.query_mode = "file"

            self.file_query_panel.show()
            self.manual_query_panel.hide()

            self._log("已切换为文件导入模式。")

        else:

            self.query_mode = "manual"

            self.file_query_panel.hide()
            self.manual_query_panel.show()

            self._log("已切换为手动输入 TITLE 模式。")

    # --------------------------------------------------------
    # 获取手动输入内容
    # --------------------------------------------------------

    def get_manual_queries(self):

        return clean_queries(
            self.manual_titles.toPlainText().splitlines()
        )

    # --------------------------------------------------------
    # 更新手动输入数量
    # --------------------------------------------------------

    def update_manual_count(self):

        self.manual_count.setText(
            f"当前输入 {len(self.get_manual_queries())} 个有效 TITLE"
        )

    # --------------------------------------------------------
    # 清空手动输入
    # --------------------------------------------------------

    def clear_manual_titles(self):

        if not self._ensure_idle():
            return

        answer = QMessageBox.question(
            self,
            "确认清空",
            "确定清空当前手动输入的全部 TITLE 吗？",
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No
        )

        if answer == QMessageBox.StandardButton.Yes:

            self.manual_titles.clear()

            self._log("手动输入内容已清空。")

    # --------------------------------------------------------
    # 选择输出路径
    # --------------------------------------------------------

    def choose_output_path(self):

        if not self._ensure_idle():
            return

        name = (
            self.output_name_edit.text().strip()
            or "TITLE整列提取结果.csv"
        )

        if not name.lower().endswith((".csv", ".xlsx")):
            name += ".csv"

        path, _ = QFileDialog.getSaveFileName(
            self,
            "选择提取结果保存位置",
            name,
            "CSV 文件 (*.csv);;Excel 文件 (*.xlsx)"
        )

        if path:

            if not path.lower().endswith((".csv", ".xlsx")):
                path += ".csv"

            self.output_path_edit.setText(path)

            self.output_name_edit.setText(
                os.path.basename(path)
            )

    # --------------------------------------------------------
    # 开始批量提取
    # --------------------------------------------------------

    def start_extract(self):

        if not self._ensure_idle():
            return

        if not self.data_files:

            QMessageBox.warning(
                self,
                "缺少原始文件",
                "请先添加原始数据文件。"
            )

            return

        # 获取查询内容

        if self.query_mode == "file":

            if not self.query_file:

                QMessageBox.warning(
                    self,
                    "缺少查询清单",
                    "请先选择 TITLE 查询清单。"
                )

                return

            if not os.path.isfile(self.query_file):

                QMessageBox.warning(
                    self,
                    "文件不存在",
                    "查询清单已经不存在，请重新选择。"
                )

                return

            try:

                queries = read_query_file(
                    self.query_file,
                    self.skip_header.isChecked()
                )

            except Exception as exc:

                QMessageBox.critical(
                    self,
                    "读取查询清单失败",
                    str(exc)
                )

                return

        else:

            queries = self.get_manual_queries()

        if not queries:

            QMessageBox.warning(
                self,
                "查询内容为空",
                "请提供至少一个有效的 TITLE。"
            )

            return

        # 获取输出路径

        output_path = self.output_path_edit.text().strip()

        if not output_path:

            suggested = (
                self.output_name_edit.text().strip()
                or "TITLE整列提取结果.csv"
            )

            if not suggested.lower().endswith((".csv", ".xlsx")):
                suggested += ".csv"

            output_path, _ = QFileDialog.getSaveFileName(
                self,
                "选择提取结果保存位置",
                suggested,
                "CSV 文件 (*.csv);;Excel 文件 (*.xlsx)"
            )

            if not output_path:
                return

            if not output_path.lower().endswith((".csv", ".xlsx")):
                output_path += ".csv"

            self.output_path_edit.setText(output_path)

            self.output_name_edit.setText(
                os.path.basename(output_path)
            )

        if not output_path.lower().endswith((".csv", ".xlsx")):

            output_path += ".csv"

            self.output_path_edit.setText(output_path)

        # 防止覆盖原始文件或查询清单

        output_real = os.path.realpath(output_path)

        protected = list(self.data_files)

        if self.query_mode == "file":
            protected.append(self.query_file)

        if any(
            os.path.realpath(path) == output_real
            for path in protected
        ):

            QMessageBox.critical(
                self,
                "路径冲突",
                "输出文件不能覆盖原始数据文件或查询清单。"
            )

            return

        # 开始处理

        self._log("=" * 55)
        self._log("开始批量提取")

        self._log(
            f"原始文件数量：{len(self.data_files)}"
        )

        self._log(
            f"查询 TITLE 数量：{len(queries)}"
        )

        self._log(
            "输出文件路径：" + output_path
        )

        self.status_label.setText(
            "运行状态：正在处理，请稍候……"
        )

        self.progress_bar.setValue(0)

        self.run_btn.setEnabled(False)

        self._set_controls_enabled(False)

        self.worker = ExtractWorker(
            self.data_files,
            queries,
            output_path
        )

        self.worker.progress.connect(
            self.on_progress
        )

        self.worker.succeeded.connect(
            self.on_success
        )

        self.worker.failed.connect(
            self.on_failure
        )

        self.worker.start()

    # --------------------------------------------------------
    # 运行期间禁用设置
    # --------------------------------------------------------

    def _set_controls_enabled(self, enabled):

        for widget in (
            self.add_data_btn,
            self.remove_data_btn,
            self.clear_data_btn,
            self.choose_query_btn,
            self.file_mode_radio,
            self.manual_mode_radio,
            self.skip_header,
            self.manual_titles,
            self.clear_manual_btn,
            self.output_name_edit,
            self.output_path_edit,
            self.browse_output_btn
        ):

            widget.setEnabled(enabled)

    # --------------------------------------------------------
    # 进度更新
    # --------------------------------------------------------

    def on_progress(self, current, total, message):

        self._log(message)

        self.progress_bar.setValue(
            int(current / max(total, 1) * 90)
        )

        self.status_label.setText(
            f"运行状态：正在处理第 {current}/{total} 个文件"
        )

    # --------------------------------------------------------
    # 提取成功
    # --------------------------------------------------------

    def on_success(self, result):

        self.progress_bar.setValue(100)

        self.status_label.setText(
            "运行状态：提取完成"
        )

        self.output_path_edit.setText(
            result["output_path"]
        )

        self._log("提取完成。")

        self._log(
            f"查询 TITLE 数量：{result['query_count']}"
        )

        self._log(
            f"匹配文件数量：{result['matched_files']}"
        )

        self._log(
            f"匹配列数量：{result['matched_columns']}"
        )

        self._log(
            f"实际输出数据行数：{result['output_rows']}"
        )

        self._log(
            "输出文件路径：" + result["output_path"]
        )

        if result["missing"]:

            self._log("以下 TITLE 未找到：")

            for title in result["missing"]:
                self._log("未找到：" + title)

        else:

            self._log(
                "所有查询 TITLE 均已找到。"
            )

        if result["errors"]:

            self._log("以下文件读取失败：")

            for error in result["errors"]:
                self._log(error)

        self.run_btn.setEnabled(True)

        self._set_controls_enabled(True)

        summary = (
            f"提取完成！\n\n"
            f"查询 TITLE：{result['query_count']} 个\n"
            f"匹配文件：{result['matched_files']} 个\n"
            f"匹配列数：{result['matched_columns']} 列\n"
            f"输出数据行数：{result['output_rows']} 行\n\n"
            f"保存位置：\n{result['output_path']}"
        )

        QMessageBox.information(
            self,
            "提取完成",
            summary
        )

    # --------------------------------------------------------
    # 提取失败
    # --------------------------------------------------------

    def on_failure(self, message):

        self.status_label.setText(
            "运行状态：提取失败"
        )

        self._log("错误：" + message)

        self.run_btn.setEnabled(True)

        self._set_controls_enabled(True)

        QMessageBox.critical(
            self,
            "提取失败",
            message
        )

    # --------------------------------------------------------
    # 关闭窗口
    # --------------------------------------------------------

    def closeEvent(self, event):

        if self.worker and self.worker.isRunning():

            answer = QMessageBox.question(
                self,
                "正在处理",
                "程序正在提取数据，确定要退出吗？"
                "退出可能导致输出文件不完整。",
                QMessageBox.StandardButton.Yes
                | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )

            if answer != QMessageBox.StandardButton.Yes:

                event.ignore()
                return

            self.worker.wait()

        event.accept()


# ============================================================
# 8. 程序入口
# ============================================================

def main():

    app = QApplication([])

    app.setApplicationName(
        "TITLE 超大型文件提取工作台"
    )

    app.setFont(
        QFont("PingFang SC", 13)
    )

    window = MainWindow()

    window.showMaximized()

    app.exec()


if __name__ == "__main__":
    main()