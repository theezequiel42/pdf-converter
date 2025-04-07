# pdf_converter_plus.py
import sys
import os
import tempfile
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QVBoxLayout,
    QPushButton, QLabel, QTextEdit, QInputDialog
)
from PyQt6.QtGui import QFont
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image
import fitz
from pdf2docx import Converter


class PDFToolApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("📄 Conversor e Utilitários PDF")
        self.setGeometry(300, 150, 600, 650)
        self.initUI()
        self.set_light_theme()

    def initUI(self):
        layout = QVBoxLayout()

        title = QLabel("🧰 Ferramentas PDF, Word e Imagem")
        title.setFont(QFont("Arial", 16))
        layout.addWidget(title)

        subtitle = QLabel("Escolha uma das opções abaixo:")
        subtitle.setFont(QFont("Arial", 12))
        layout.addWidget(subtitle)

        buttons = [
            ("✂️ Dividir PDF", self.dividir_pdf),
            ("🖼️ Imagens para PDF", self.imagens_para_pdf),
            ("🧾 PDF para Imagens", self.pdf_para_imagens),
            ("📝 PDF para Word", self.pdf_para_word),
            ("📎 Mesclar PDFs", self.mesclar_pdfs),
            ("🔒 Proteger PDF com Senha", self.proteger_pdf),
            ("📉 Comprimir PDF", self.comprimir_pdf),
        ]

        for label, action in buttons:
            btn = QPushButton(label)
            btn.setMinimumHeight(40)
            btn.clicked.connect(action)
            layout.addWidget(btn)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setFont(QFont("Courier", 10))
        layout.addWidget(self.log)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def set_light_theme(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #ffffff; }
            QPushButton {
                background-color: #4a90e2;
                color: white;
                font-size: 14px;
                padding: 10px;
                border-radius: 6px;
            }
            QPushButton:hover { background-color: #357ab8; }
            QTextEdit {
                background-color: #f7f7f7;
                color: #000000;
                border: 1px solid #999;
                border-radius: 5px;
            }
            QLabel {
                font-weight: bold;
                margin-bottom: 5px;
            }
        """)

    def log_msg(self, msg):
        self.log.append(msg)

    def dividir_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Escolher PDF", "", "PDF Files (*.pdf)")
        if not path:
            return
        paginas, ok = QInputDialog.getText(self, "Selecionar Páginas", "Digite páginas (ex: 1,2,4-6):")
        if not ok:
            return
        try:
            reader = PdfReader(path)
            writer = PdfWriter()
            total = len(reader.pages)

            paginas_selecionadas = []
            for parte in paginas.split(","):
                if "-" in parte:
                    inicio, fim = map(int, parte.split("-"))
                    paginas_selecionadas.extend(range(inicio - 1, fim))
                else:
                    paginas_selecionadas.append(int(parte) - 1)

            paginas_selecionadas = sorted(set(p for p in paginas_selecionadas if 0 <= p < total))
            for p in paginas_selecionadas:
                writer.add_page(reader.pages[p])

            output_path, _ = QFileDialog.getSaveFileName(self, "Salvar PDF dividido como...", "dividido.pdf", "PDF Files (*.pdf)")
            if output_path:
                with open(output_path, "wb") as f:
                    writer.write(f)
                self.log_msg(f"✅ PDF salvo em {output_path}")
                os.startfile(output_path)

        except Exception as e:
            self.log_msg(f"❌ Erro: {e}")

    def imagens_para_pdf(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar Imagens", "", "Imagens (*.png *.jpg *.jpeg)")
        if not files:
            return
        try:
            images = [Image.open(p).convert("RGB") for p in files]
            output_path, _ = QFileDialog.getSaveFileName(self, "Salvar PDF como...", "imagens.pdf", "PDF Files (*.pdf)")
            if output_path:
                images[0].save(output_path, save_all=True, append_images=images[1:])
                self.log_msg(f"✅ PDF salvo em {output_path}")
                os.startfile(output_path)
        except Exception as e:
            self.log_msg(f"❌ Erro: {e}")

    def pdf_para_imagens(self):
        path, _ = QFileDialog.getOpenFileName(self, "Escolher PDF", "", "PDF Files (*.pdf)")
        if not path:
            return
        try:
            doc = fitz.open(path)
            dir_path = QFileDialog.getExistingDirectory(self, "Escolher pasta para salvar as imagens")
            if dir_path:
                for i, page in enumerate(doc):
                    pix = page.get_pixmap()
                    img_path = os.path.join(dir_path, f"pagina_{i + 1}.png")
                    pix.save(img_path)
                    self.log_msg(f"📸 Página {i + 1} salva em {img_path}")
                os.startfile(dir_path)
        except Exception as e:
            self.log_msg(f"❌ Erro: {e}")

    def pdf_para_word(self):
        path, _ = QFileDialog.getOpenFileName(self, "Escolher PDF", "", "PDF Files (*.pdf)")
        if not path:
            return
        try:
            output_path, _ = QFileDialog.getSaveFileName(self, "Salvar Word como...", "convertido.docx", "Word Files (*.docx)")
            if output_path:
                cv = Converter(path)
                cv.convert(output_path, start=0, end=None)
                cv.close()
                self.log_msg(f"✅ Word salvo em {output_path}")
                os.startfile(output_path)
        except Exception as e:
            self.log_msg(f"❌ Erro: {e}")

    def mesclar_pdfs(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Selecionar PDFs para Mesclar", "", "PDF Files (*.pdf)")
        if not files or len(files) < 2:
            self.log_msg("⚠️ Selecione pelo menos dois arquivos.")
            return
        try:
            writer = PdfWriter()
            for file in files:
                reader = PdfReader(file)
                for page in reader.pages:
                    writer.add_page(page)

            output_path, _ = QFileDialog.getSaveFileName(self, "Salvar PDF mesclado como...", "mesclado.pdf", "PDF Files (*.pdf)")
            if output_path:
                with open(output_path, "wb") as f:
                    writer.write(f)
                self.log_msg(f"✅ PDF mesclado salvo em {output_path}")
                os.startfile(output_path)

        except Exception as e:
            self.log_msg(f"❌ Erro: {e}")

    def proteger_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Escolher PDF", "", "PDF Files (*.pdf)")
        if not path:
            return
        senha, ok = QInputDialog.getText(self, "Definir Senha", "Digite uma senha para proteger o PDF:")
        if not ok or not senha:
            return
        try:
            reader = PdfReader(path)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            writer.encrypt(senha)

            output_path, _ = QFileDialog.getSaveFileName(self, "Salvar PDF protegido como...", "protegido.pdf", "PDF Files (*.pdf)")
            if output_path:
                with open(output_path, "wb") as f:
                    writer.write(f)
                self.log_msg(f"🔒 PDF protegido salvo em {output_path}")
                os.startfile(output_path)

        except Exception as e:
            self.log_msg(f"❌ Erro: {e}")

    def comprimir_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Escolher PDF para Comprimir", "", "PDF Files (*.pdf)")
        if not path:
            return
        try:
            doc = fitz.open(path)
            output_path, _ = QFileDialog.getSaveFileName(self, "Salvar PDF comprimido como...", "comprimido.pdf", "PDF Files (*.pdf)")
            if output_path:
                new_doc = fitz.open()
                for page in doc:
                    pix = page.get_pixmap(dpi=72)  # Reduz DPI
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    img_path = os.path.join(tempfile.gettempdir(), f"{os.urandom(4).hex()}.jpg")
                    img.save(img_path, "JPEG", quality=50)

                    rect = fitz.Rect(0, 0, pix.width, pix.height)
                    page_img = new_doc.new_page(width=pix.width, height=pix.height)
                    page_img.insert_image(rect, filename=img_path)

                new_doc.save(output_path)
                new_doc.close()
                self.log_msg(f"📉 PDF comprimido salvo em {output_path}")
                os.startfile(output_path)

        except Exception as e:
            self.log_msg(f"❌ Erro: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFToolApp()
    window.show()
    sys.exit(app.exec())
