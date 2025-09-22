import sys
import os
import chardet
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel,
    QMessageBox, QTextEdit, QPushButton, QHBoxLayout,
    QTabWidget, QListWidget
)
from PyQt6.QtCore import Qt
from opencc import OpenCC
from docx import Document
import subprocess

# -------------------- 文件 Tab --------------------
class FileTab(QWidget):
    def __init__(self, parent_app):
        super().__init__()
        self.parent_app = parent_app
        self.setAcceptDrops(True)

        layout = QVBoxLayout()
        self.setLayout(layout)

        self.label_file = QLabel("将文件或文件夹拖入此窗口\n支持 TXT、MD、HTML、DOCX 文本文件")
        self.label_file.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label_file)

        self.file_list = QListWidget()
        layout.addWidget(self.file_list)

        btn_layout = QHBoxLayout()
        self.open_file_btn = QPushButton("打开文件")
        self.open_file_btn.clicked.connect(self.open_selected_file)
        self.open_location_btn = QPushButton("打开文件位置")
        self.open_location_btn.clicked.connect(self.open_file_location)
        btn_layout.addWidget(self.open_file_btn)
        btn_layout.addWidget(self.open_location_btn)
        layout.addLayout(btn_layout)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            self.parent_app.process_path(path, self.file_list)
        self.label_file.setText("文件转换完成！")
        QMessageBox.information(self.parent_app, "完成", "所有文件已转换完成！")

    def open_selected_file(self):
        item = self.file_list.currentItem()
        if item and os.path.exists(item.text()):
            os.startfile(item.text())
        else:
            QMessageBox.warning(self, "错误", "文件不存在！")

    def open_file_location(self):
        item = self.file_list.currentItem()
        if item:
            folder = os.path.dirname(item.text())
            if os.path.exists(folder):
                os.startfile(folder)
            else:
                QMessageBox.warning(self, "错误", "文件夹不存在！")

# -------------------- 主程序 --------------------
class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("简体→台湾繁体转换器")
        self.resize(900, 600)

        self.cc = OpenCC('s2twp')
        self.log = []

        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # 实时转换 Tab
        self.tab_realtime = QWidget()
        self.tabs.addTab(self.tab_realtime, "实时转换")
        self.init_realtime_tab()

        # 文件转换 Tab
        self.tab_file = FileTab(self)
        self.tabs.addTab(self.tab_file, "文件转换")

    # -------------------- 实时转换 --------------------
    def init_realtime_tab(self):
        layout = QVBoxLayout()
        self.tab_realtime.setLayout(layout)

        text_layout = QHBoxLayout()

        # 左侧简体输入
        left_layout = QVBoxLayout()
        self.input_text = QTextEdit()
        self.input_text.setPlaceholderText("在这里输入简体中文...")
        self.copy_input_btn = QPushButton("复制简体中文")
        self.copy_input_btn.clicked.connect(self.copy_input_text)
        left_layout.addWidget(self.input_text)
        left_layout.addWidget(self.copy_input_btn)
        text_layout.addLayout(left_layout)

        # 右侧繁体输出
        right_layout = QVBoxLayout()
        self.output_text = QTextEdit()
        self.output_text.setPlaceholderText("台湾繁体显示...")
        self.output_text.setReadOnly(True)
        self.copy_output_btn = QPushButton("复制繁体中文")
        self.copy_output_btn.clicked.connect(self.copy_output_text)
        right_layout.addWidget(self.output_text)
        right_layout.addWidget(self.copy_output_btn)
        text_layout.addLayout(right_layout)

        layout.addLayout(text_layout)

        # 滚轮同步
        self.input_text.verticalScrollBar().valueChanged.connect(
            lambda value: self.output_text.verticalScrollBar().setValue(value)
        )
        self.output_text.verticalScrollBar().valueChanged.connect(
            lambda value: self.input_text.verticalScrollBar().setValue(value)
        )

        # 实时转换信号
        self.input_text.textChanged.connect(self.update_conversion)

    def update_conversion(self):
        text = self.input_text.toPlainText()
        converted = self.cc.convert(text)

        highlighted = ""
        for i, s_char in enumerate(text):
            t_char = converted[i] if i < len(converted) else ""
            if s_char != t_char and t_char != "":
                if s_char in [" ", "\t"]:
                    highlighted += s_char
                elif s_char == "\n":
                    highlighted += "<br>"
                else:
                    highlighted += f'<span style="background-color:yellow">{t_char}</span>'
            else:
                if s_char == "\n":
                    highlighted += "<br>"
                elif s_char == "\t":
                    highlighted += "&nbsp;&nbsp;&nbsp;&nbsp;"
                elif s_char == " ":
                    highlighted += "&nbsp;"
                else:
                    highlighted += s_char

        # 剩余字符高亮
        if len(converted) > len(text):
            for t_char in converted[len(text):]:
                if t_char == "\n":
                    highlighted += "<br>"
                elif t_char == "\t":
                    highlighted += "&nbsp;&nbsp;&nbsp;&nbsp;"
                elif t_char == " ":
                    highlighted += "&nbsp;"
                else:
                    highlighted += f'<span style="background-color:yellow">{t_char}</span>'

        self.output_text.setHtml(highlighted)

    def copy_input_text(self):
        QApplication.clipboard().setText(self.input_text.toPlainText())
        QMessageBox.information(self, "提示", "简体文本已复制到剪贴板！")

    def copy_output_text(self):
        QApplication.clipboard().setText(self.output_text.toPlainText())
        QMessageBox.information(self, "提示", "繁体文本已复制到剪贴板！")

    # -------------------- 文件转换 --------------------
    def process_path(self, path, file_list_widget):
        if os.path.isfile(path):
            new_file = self.convert_file(path)
            if new_file:
                file_list_widget.addItem(new_file)
        elif os.path.isdir(path):
            for root, dirs, files in os.walk(path):
                for file in files:
                    if file.lower().endswith(('.txt', '.md', '.html', '.htm', '.docx')):
                        full_path = os.path.join(root, file)
                        new_file = self.convert_file(full_path)
                        if new_file:
                            file_list_widget.addItem(new_file)

    def convert_file(self, filepath):
        try:
            ext = os.path.splitext(filepath)[1].lower()
            if ext in ('.txt', '.md', '.html', '.htm'):
                with open(filepath, 'rb') as f:
                    raw = f.read()
                detected = chardet.detect(raw)
                encoding = detected['encoding'] or 'utf-8'
                text = raw.decode(encoding, errors='ignore')
                converted = self.cc.convert(text)
                new_file = f"{os.path.splitext(filepath)[0]}_tw{ext}"
                with open(new_file, 'w', encoding='utf-8') as f:
                    f.write(converted)
            elif ext == '.docx':
                doc = Document(filepath)
                for p in doc.paragraphs:
                    p.text = self.cc.convert(p.text)
                new_file = f"{os.path.splitext(filepath)[0]}_tw.docx"
                doc.save(new_file)
            self.log.append(f"转换成功: {filepath} -> {new_file}")
            return new_file
        except Exception as e:
            self.log.append(f"转换失败: {filepath}, 原因: {e}")
            return None

    def save_log(self, log_file='conversion_log.txt'):
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(self.log))

# -------------------- 主程序入口 --------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ConverterApp()
    window.show()
    exit_code = app.exec()
    window.save_log()
    sys.exit(exit_code)
