
import sys
import os
import csv
import json
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QGridLayout,
    QVBoxLayout,
)
from PySide6.QtCore import Qt

from openpyxl import load_workbook, Workbook


# ============================================================
# 配置文件
# ============================================================

CONFIG_FILE = os.path.expanduser("~/.fai_mapping_config.json")


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)

                # 兼容旧版本配置文件
                if "output_dir" not in cfg:
                    old_output = cfg.get("output_path", "")
                    cfg["output_dir"] = (
                        os.path.dirname(old_output)
                        if old_output else ""
                    )

                return cfg

        except Exception:
            pass

    return {
        "template_path": "",
        "title_path": "",
        "output_dir": ""
    }


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ============================================================
# Excel列号转换
# ============================================================

def col_num_to_letter(n):
    """数字列号转Excel列字母，1->A, 26->Z, 27->AA"""

    if n is None:
        return None

    s = ""

    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(ord('A') + rem) + s

    return s


# ============================================================
# FAI对应关系生成核心程序
# ============================================================

def run_mapping(template_path, title_path, output_path):

    # ========= 1.读取Template CPK工作表 =========

    wb_temp = load_workbook(
        template_path,
        read_only=True,
        data_only=True
    )

    if "CPK" not in wb_temp.sheetnames:
        wb_temp.close()
        raise Exception("Template文件中未找到CPK工作表！")

    ws_cpk = wb_temp["CPK"]

    all_rows = list(ws_cpk.iter_rows(values_only=True))

    wb_temp.close()

    # ========= 扫描前20行查找FAI#ROW表头 =========

    fai_col_idx = None
    start_data_row_idx = None

    for scan_idx, row_tuple in enumerate(all_rows[:20], start=1):

        for col_idx, cell_val in enumerate(row_tuple):

            if cell_val is not None:

                s_val = str(cell_val).strip()

                if s_val == "FAI#ROW":

                    fai_col_idx = col_idx
                    start_data_row_idx = scan_idx

                    break

        if fai_col_idx is not None:
            break

    if fai_col_idx is None:
        raise Exception(
            "在CPK工作表前20行未找到'FAI#ROW'表头，请检查文件！"
        )

    # ========= 固定第9行映射到A列 =========

    if len(all_rows) < 9:
        raise Exception("Template文件的CPK工作表不足9行！")

    row9_tuple = all_rows[8]

    val_row9 = (
        row9_tuple[fai_col_idx]
        if fai_col_idx < len(row9_tuple)
        else None
    )

    fai_text_row9 = (
        str(val_row9)
        if val_row9 is not None
        else ""
    )

    fixed_item = {
        "template_FAI#": fai_text_row9,
        "template_excel_row": 9,
        "title_column_letter": "A",
        "title_column_index": 1
    }

    # ========= 遍历表头下一行开始的全部数据 =========

    fai_list = []

    for mem_idx in range(start_data_row_idx, len(all_rows)):

        excel_physical_row = mem_idx + 1

        row_tuple = all_rows[mem_idx]

        val = (
            row_tuple[fai_col_idx]
            if fai_col_idx < len(row_tuple)
            else None
        )

        if val is None:
            fai_text = ""
        else:
            fai_text = str(val)

        # 过滤空FAI#ROW
        if fai_text.strip() == "":
            continue

        fai_list.append((fai_text, excel_physical_row))

    # ========= 2.只读取Title.csv表头 =========

    with open(
        title_path,
        "r",
        encoding="utf-8",
        newline=""
    ) as f:

        reader = csv.reader(f)

        title_header = next(reader)

    # ========= 3.逐行匹配 =========

    result_rows = []

    for fai_text, excel_row in fai_list:

        col_idx_1base = None

        if fai_text.strip() != "":

            try:
                col_idx_1base = title_header.index(fai_text) + 1

            except ValueError:
                col_idx_1base = None

        col_letter = col_num_to_letter(col_idx_1base)

        result_rows.append({
            "template_FAI#": fai_text,
            "template_excel_row": excel_row,
            "title_column_letter": col_letter,
            "title_column_index": col_idx_1base
        })

    full_sheet1 = [fixed_item] + result_rows

    # ========= 4.输出对应关系表 =========

    wb_out = Workbook()

    ws1 = wb_out.active

    ws1.title = "对应关系表"

    headers = [
        "template_FAI#",
        "template_excel_row",
        "title_column_letter",
        "title_column_index"
    ]

    ws1.append(headers)

    for item in full_sheet1:

        ws1.append([
            item["template_FAI#"],
            item["template_excel_row"],
            item["title_column_letter"],
            item["title_column_index"]
        ])

    wb_out.save(output_path)

    return len(full_sheet1)


# ============================================================
# 自动生成输出文件名
# ============================================================

def generate_output_path(output_dir):

    # 获取当前时间，精确到分钟
    current_time = datetime.now().strftime("%Y%m%d%H%M")

    # 自动生成文件名
    filename = f"对应关系表{current_time}.xlsx"

    # 拼接输出路径
    output_path = os.path.join(output_dir, filename)

    return output_path


# ============================================================
# PySide6 GUI界面
# ============================================================

class App(QWidget):

    def __init__(self, cfg):

        super().__init__()

        self.cfg = cfg

        self.init_ui()

    # ========================================================
    # 初始化界面
    # ========================================================

    def init_ui(self):

        self.setWindowTitle(
            "FAI对应关系生成工具｜过滤空行｜FAI#ROW｜第9行固定A列"
        )

        self.resize(850, 310)

        self.setMinimumWidth(750)

        main_layout = QVBoxLayout()

        main_layout.setSpacing(18)

        main_layout.setContentsMargins(22, 22, 22, 22)

        grid = QGridLayout()

        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(18)

        # ====================================================
        # Template文件
        # ====================================================

        label_template = QLabel("Template文件（xlsx）")

        self.input_template = QLineEdit()

        self.input_template.setText(
            self.cfg.get("template_path", "")
        )

        self.input_template.setPlaceholderText(
            "请选择Template Excel文件"
        )

        btn_template = QPushButton("选择文件")

        btn_template.clicked.connect(
            self.select_template
        )

        grid.addWidget(label_template, 0, 0)
        grid.addWidget(self.input_template, 0, 1)
        grid.addWidget(btn_template, 0, 2)

        # ====================================================
        # Title文件
        # ====================================================

        label_title = QLabel("Title文件（csv）")

        self.input_title = QLineEdit()

        self.input_title.setText(
            self.cfg.get("title_path", "")
        )

        self.input_title.setPlaceholderText(
            "请选择Title CSV文件"
        )

        btn_title = QPushButton("选择文件")

        btn_title.clicked.connect(
            self.select_title
        )

        grid.addWidget(label_title, 1, 0)
        grid.addWidget(self.input_title, 1, 1)
        grid.addWidget(btn_title, 1, 2)

        # ====================================================
        # 输出文件夹
        # ====================================================

        label_output = QLabel("输出文件夹")

        self.input_output = QLineEdit()

        self.input_output.setText(
            self.cfg.get("output_dir", "")
        )

        self.input_output.setPlaceholderText(
            "请选择输出文件夹，文件名将自动生成"
        )

        btn_output = QPushButton("选择文件夹")

        btn_output.clicked.connect(
            self.select_output
        )

        grid.addWidget(label_output, 2, 0)
        grid.addWidget(self.input_output, 2, 1)
        grid.addWidget(btn_output, 2, 2)

        main_layout.addLayout(grid)

        # ====================================================
        # 自动命名说明
        # ====================================================

        self.label_filename = QLabel(
            "输出文件名：对应关系表 + 当前处理时间"
        )

        self.label_filename.setWordWrap(True)

        self.label_filename.setStyleSheet(
            """
            QLabel {
                color: #3478d4;
                font-size: 12px;
                padding: 9px;
                background-color: #f0f5fc;
                border-radius: 5px;
            }
            """
        )

        main_layout.addWidget(self.label_filename)

        # ====================================================
        # 功能说明
        # ====================================================

        self.label_info = QLabel(
            "处理规则：扫描CPK前20行查找FAI#ROW，"
            "第9行固定映射A列，自动过滤空FAI#ROW。"
        )

        self.label_info.setWordWrap(True)

        self.label_info.setStyleSheet(
            """
            QLabel {
                color: #777777;
                font-size: 12px;
                padding: 8px;
                background-color: #f3f4f6;
                border-radius: 5px;
            }
            """
        )

        main_layout.addWidget(self.label_info)

        # ====================================================
        # 生成按钮
        # ====================================================

        self.btn_generate = QPushButton("生成对应关系表")

        self.btn_generate.setFixedHeight(42)

        self.btn_generate.setCursor(
            Qt.CursorShape.PointingHandCursor
        )

        self.btn_generate.setStyleSheet(
            """
            QPushButton {
                background-color: #3478d4;
                color: white;
                font-size: 14px;
                font-weight: bold;
                border: none;
                border-radius: 6px;
            }

            QPushButton:hover {
                background-color: #2465b9;
            }

            QPushButton:pressed {
                background-color: #1c559f;
            }

            QPushButton:disabled {
                background-color: #a0a0a0;
            }
            """
        )

        self.btn_generate.clicked.connect(
            self.do_generate
        )

        main_layout.addWidget(self.btn_generate)

        self.setLayout(main_layout)

        # ====================================================
        # 界面样式
        # ====================================================

        self.setStyleSheet(
            """
            QWidget {
                background-color: #ffffff;
                font-family: Arial, "PingFang SC";
                font-size: 13px;
            }

            QLabel {
                color: #333333;
            }

            QLineEdit {
                border: 1px solid #cccccc;
                border-radius: 5px;
                padding: 7px;
                background-color: white;
                min-height: 22px;
            }

            QLineEdit:focus {
                border: 1px solid #3478d4;
            }

            QPushButton {
                padding: 6px 15px;
                border: 1px solid #cccccc;
                border-radius: 5px;
                background-color: #f7f7f7;
            }

            QPushButton:hover {
                background-color: #eeeeee;
            }
            """
        )

    # ========================================================
    # 选择Template文件
    # ========================================================

    def select_template(self):

        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Template文件",
            "",
            "Excel Files (*.xlsx)"
        )

        if path:
            self.input_template.setText(path)

    # ========================================================
    # 选择Title文件
    # ========================================================

    def select_title(self):

        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Title文件",
            "",
            "CSV Files (*.csv)"
        )

        if path:
            self.input_title.setText(path)

    # ========================================================
    # 选择输出文件夹
    # ========================================================

    def select_output(self):

        current_dir = self.input_output.text().strip()

        if not os.path.isdir(current_dir):
            current_dir = ""

        path = QFileDialog.getExistingDirectory(
            self,
            "选择输出文件夹",
            current_dir
        )

        if path:
            self.input_output.setText(path)

    # ========================================================
    # 执行生成
    # ========================================================

    def do_generate(self):

        tpl_path = self.input_template.text().strip()

        title_path = self.input_title.text().strip()

        output_dir = self.input_output.text().strip()

        # ========= 检查Template路径 =========

        if not tpl_path:

            QMessageBox.critical(
                self,
                "错误",
                "请选择Template文件！"
            )

            return

        if not os.path.exists(tpl_path):

            QMessageBox.critical(
                self,
                "错误",
                "Template文件不存在！"
            )

            return

        # ========= 检查Title路径 =========

        if not title_path:

            QMessageBox.critical(
                self,
                "错误",
                "请选择Title文件！"
            )

            return

        if not os.path.exists(title_path):

            QMessageBox.critical(
                self,
                "错误",
                "Title文件不存在！"
            )

            return

        # ========= 检查输出文件夹 =========

        if not output_dir:

            QMessageBox.critical(
                self,
                "错误",
                "请选择输出文件夹！"
            )

            return

        if not os.path.isdir(output_dir):

            QMessageBox.critical(
                self,
                "错误",
                "输出文件夹不存在！"
            )

            return

        # ========= 自动生成输出文件名 =========

        out_path = generate_output_path(output_dir)

        # ========= 保存配置 =========

        new_cfg = {
            "template_path": tpl_path,
            "title_path": title_path,
            "output_dir": output_dir
        }

        try:

            save_config(new_cfg)

        except Exception as e:

            QMessageBox.warning(
                self,
                "配置保存提醒",
                f"配置保存失败，但仍将尝试执行：\n{str(e)}"
            )

        # ========= 执行程序 =========

        self.btn_generate.setEnabled(False)

        self.btn_generate.setText("正在生成，请稍候...")

        QApplication.processEvents()

        try:

            total = run_mapping(
                tpl_path,
                title_path,
                out_path
            )

            QMessageBox.information(
                self,
                "完成",
                f"处理完成！\n\n"
                f"共输出 {total} 行记录\n\n"
                f"包含第9行固定A列映射\n"
                f"已过滤空FAI#ROW行\n\n"
                f"自动生成的文件名：\n"
                f"{os.path.basename(out_path)}\n\n"
                f"输出位置：\n{out_path}"
            )

        except Exception as e:

            QMessageBox.critical(
                self,
                "程序异常",
                f"执行失败：\n{str(e)}"
            )

        finally:

            self.btn_generate.setEnabled(True)

            self.btn_generate.setText("生成对应关系表")


# ============================================================
# 主程序
# ============================================================

def main():

    cfg = load_config()

    # ========= 无界面模式 =========

    if len(sys.argv) == 5 and sys.argv[1] == "--headless":

        tpl = sys.argv[2]

        tit = sys.argv[3]

        out = sys.argv[4]

        total = run_mapping(tpl, tit, out)

        print(
            f"headless模式执行完毕，输出{total}行"
            f"（包含第9行固定A列映射，已过滤空FAI#ROW行）"
        )

        return

    # ========= PySide6 GUI模式 =========

    app = QApplication(sys.argv)

    app.setStyle("Fusion")

    window = App(cfg)

    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":

    main()
