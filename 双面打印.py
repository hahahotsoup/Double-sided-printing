import sys
import PyPDF2
import os
from docx2pdf import convert
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QProgressBar

class PDFSplitter(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('PDF 工具')
        self.setGeometry(100, 100, 300, 200)

        self.label = QLabel('选择一个工具:', self)
        self.split_button = QPushButton('拆解PDF', self)
        self.split_button.clicked.connect(self.choosePDF)
        self.merge_button = QPushButton('合并PDF', self)
        self.merge_button.clicked.connect(self.mergePDF)
        self.convert_button = QPushButton('DOCX转PDF', self)
        self.convert_button.clicked.connect(self.convertDOCXtoPDF)
        self.progress = QProgressBar(self)
        self.progress.setVisible(False)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.split_button)
        layout.addWidget(self.merge_button)
        layout.addWidget(self.convert_button)
        layout.addWidget(self.progress)
        self.setLayout(layout)

    def choosePDF(self):
        options = QFileDialog.Options()
        file_paths, _ = QFileDialog.getOpenFileNames(self, "选择PDF文件", "", "PDF 文件 (*.pdf)", options=options)
        if file_paths:
            output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录")
            if output_dir:
                self.progress.setVisible(True)
                self.progress.setMaximum(len(file_paths))
                for i, file_path in enumerate(file_paths):
                    self.splitPDF(file_path, output_dir)
                    self.progress.setValue(i + 1)
                self.progress.setVisible(False)
                self.label.setText('所有PDF文件已成功拆解!')

    def splitPDF(self, input_path, output_dir):
        base_filename = os.path.basename(input_path).replace('.pdf', '')
        odd_dir = os.path.join(output_dir, '奇数页')
        even_dir = os.path.join(output_dir, '偶数页')
        os.makedirs(odd_dir, exist_ok=True)
        os.makedirs(even_dir, exist_ok=True)
        output_path_odd = os.path.join(odd_dir, f"{base_filename}_奇数页.pdf")
        output_path_even = os.path.join(even_dir, f"{base_filename}_偶数页.pdf")
        
        with open(input_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            writer_odd = PyPDF2.PdfWriter()
            writer_even = PyPDF2.PdfWriter()
            
            for i in range(len(reader.pages)):
                if (i + 1) % 2 != 0:  # 奇数页
                    writer_odd.add_page(reader.pages[i])
                else:  # 偶数页
                    writer_even.add_page(reader.pages[i])
            if len(writer_odd.pages) > len(writer_even.pages):
                for _ in range(len(writer_odd.pages) - len(writer_even.pages)):
                    writer_even.add_blank_page(width=595, height=842)  # A4 大小
            
            with open(output_path_odd, 'wb') as out_odd:
                writer_odd.write(out_odd)
            with open(output_path_even, 'wb') as out_even:
                writer_even.write(out_even)

    def mergePDF(self):
        options = QFileDialog.Options()
        file_paths, _ = QFileDialog.getOpenFileNames(self, "选择要合并的PDF文件", "", "PDF 文件 (*.pdf)", options=options)
        if file_paths:
            output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录")
            if output_dir:
                merger = PyPDF2.PdfMerger()
                for file_path in file_paths:
                    merger.append(file_path)
                output_path = f"{output_dir}/合并的文件.pdf"
                with open(output_path, 'wb') as out:
                    merger.write(out)
                self.label.setText('PDF文件已成功合并!')

    def convertDOCXtoPDF(self):
        options = QFileDialog.Options()
        file_paths, _ = QFileDialog.getOpenFileNames(self, "选择DOCX文件", "", "DOCX 文件 (*.docx)", options=options)
        if file_paths:
            output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录")
            if output_dir:
                self.progress.setVisible(True)
                self.progress.setMaximum(len(file_paths))
                for i, file_path in enumerate(file_paths):
                    convert(file_path, output_dir)
                    self.progress.setValue(i + 1)
                self.progress.setVisible(False)
                self.label.setText('所有DOCX文件已成功转换为PDF!')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PDFSplitter()
    ex.show()
    sys.exit(app.exec_())
