import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QListWidget, QLabel, 
                             QFileDialog, QMessageBox, QComboBox, QSpinBox, QGroupBox, QProgressBar)
from PyQt5.QtCore import Qt
import fitz  # PyMuPDF

class InvoiceMerger(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("12306 发票合并工具__WPython")
        self.setGeometry(100, 100, 500, 700)
        
        self.files = []
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        # 文件列表
        file_group = QGroupBox("已选择的发票文件")
        file_layout = QVBoxLayout()
        
        self.list_widget = QListWidget()
        file_layout.addWidget(self.list_widget)

        btn_layout = QHBoxLayout()
        self.btn_add = QPushButton("添加发票")
        self.btn_add.clicked.connect(self.add_files)
        self.btn_clear = QPushButton("清空列表")
        self.btn_clear.clicked.connect(self.clear_files)
        
        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_clear)
        file_layout.addLayout(btn_layout)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # 合并设置
        setting_group = QGroupBox("合并设置")
        setting_layout = QHBoxLayout()
        
        setting_layout.addWidget(QLabel("每页数量:"))
        self.combo_layout = QComboBox()
        self.combo_layout.addItems(["2张/页", "4张/页"])
        setting_layout.addWidget(self.combo_layout)

        setting_layout.addWidget(QLabel("排列方式:"))
        self.combo_direction = QComboBox()
        self.combo_direction.addItems(["竖向", "横向"])
        setting_layout.addWidget(self.combo_direction)

        setting_layout.addWidget(QLabel("页边距:"))
        self.spin_margin = QSpinBox()
        self.spin_margin.setRange(0, 100)
        self.spin_margin.setValue(30)
        setting_layout.addWidget(self.spin_margin)

        setting_group.setLayout(setting_layout)
        layout.addWidget(setting_group)

        # 执行按钮
        self.btn_merge = QPushButton("开始合并并保存")
        self.btn_merge.setStyleSheet("background-color: #4CAF50; color: white; font-size: 16px; padding: 10px;")
        self.btn_merge.clicked.connect(self.merge_invoices)
        layout.addWidget(self.btn_merge)
        
        # 进度条
        self.progress = QProgressBar()
        self.progress.setFormat("就绪")
        layout.addWidget(self.progress)

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择发票文件", "", "PDF Files (*.pdf)"
        )
        if files:
            self.files.extend(files)
            self.list_widget.clear()
            for f in self.files:
                self.list_widget.addItem(os.path.basename(f))

    def clear_files(self):
        self.files = []
        self.list_widget.clear()
        self.progress.setFormat("就绪")
        self.progress.setValue(0)

    def merge_invoices(self):
        if not self.files:
            QMessageBox.warning(self, "警告", "请先添加发票文件！")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "保存合并后的PDF", "merged_invoices.pdf", "PDF Files (*.pdf)"
        )
        if not save_path:
            return

        try:
            layout_mode = self.combo_layout.currentIndex()
            direction_mode = self.combo_direction.currentIndex()
            margin = self.spin_margin.value()
            
            page_width, page_height = 595, 842
            
            self.btn_merge.setEnabled(False)
            self.progress.setFormat("正在处理...")
            self.progress.setMaximum(len(self.files))
            QApplication.processEvents()

            new_doc = fitz.open()
            success_count = 0

            # 2张模式
            if layout_mode == 0:
                if direction_mode == 0:  # 竖向 - 上下排列
                    rows, cols = 2, 1
                    rotation = 0
                else:  # 横向 - 竖向模式转270°
                    rows, cols = 2, 1
                    rotation = 270
            # 4张模式
            else:
                if direction_mode == 0:  # 竖向 - 2x2排列
                    rows, cols = 2, 2
                    rotation = 0
                else:  # 横向 - 竖向模式转270°
                    rows, cols = 2, 2
                    rotation = 270

            cell_w = (page_width - 2 * margin) / cols
            cell_h = (page_height - 2 * margin) / rows

            invoice_count_in_page = 0
            current_page = None

            for i, file_path in enumerate(self.files):
                try:
                    temp_doc = fitz.open(file_path)
                    
                    for page_idx in range(len(temp_doc)):
                        if invoice_count_in_page == 0:
                            current_page = new_doc.new_page(-1, width=page_width, height=page_height)
                        
                        r = invoice_count_in_page // cols
                        c = invoice_count_in_page % cols
                        
                        x0 = margin + c * cell_w
                        y0 = margin + r * cell_h
                        x1 = margin + (c + 1) * cell_w
                        y1 = margin + (r + 1) * cell_h
                        
                        if x1 <= x0 or y1 <= y0:
                            continue
                            
                        target_rect = fitz.Rect(x0, y0, x1, y1)
                        
                        # 缩放留白
                        scale_factor = 0.96
                        w = target_rect.width * scale_factor
                        h = target_rect.height * scale_factor
                        
                        center_x = (x0 + x1) / 2
                        center_y = (y0 + y1) / 2
                        
                        target_rect = fitz.Rect(
                            center_x - w / 2,
                            center_y - h / 2,
                            center_x + w / 2,
                            center_y + h / 2
                        )

                        current_page.show_pdf_page(target_rect, temp_doc, page_idx, rotate=rotation)
                        
                        invoice_count_in_page += 1
                        if invoice_count_in_page == rows * cols:
                            invoice_count_in_page = 0
                            
                    temp_doc.close()
                    success_count += 1
                    
                except Exception as e:
                    print(f"处理文件失败 {file_path}: {e}")
                
                self.progress.setValue(i + 1)
                self.progress.setFormat(f"处理中: {os.path.basename(file_path)}")
                QApplication.processEvents()

            new_doc.save(save_path)
            new_doc.close()
            
            QMessageBox.information(self, "成功", f"处理完成！\n成功合并 {success_count} 个文件。\n文件已保存至:\n{save_path}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"发生未知错误:\n{str(e)}")
        
        finally:
            self.btn_merge.setEnabled(True)
            self.progress.setFormat("就绪")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    font = QApplication.font()
    font.setFamily("Microsoft YaHei")
    app.setFont(font)
    
    window = InvoiceMerger()
    window.show()
    sys.exit(app.exec_())