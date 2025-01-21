import os
import subprocess
from tkinter import Tk, filedialog, simpledialog, messagebox, Button
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image

def abrir_diretorio(caminho):
    """Abre o diretório onde o arquivo foi salvo."""
    if os.name == "nt":  # Windows
        os.startfile(caminho)
    elif os.name == "posix":  # macOS ou Linux
        subprocess.run(["open" if os.uname().sysname == "Darwin" else "xdg-open", caminho])
    else:
        messagebox.showerror("Erro", "Sistema operacional não suportado para abrir o diretório.")

def dividir_pdf():
    root = Tk()
    root.withdraw()

    arquivo_pdf = filedialog.askopenfilename(
        title="Selecione o arquivo PDF",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )

    if not arquivo_pdf:
        messagebox.showwarning("Aviso", "Nenhum arquivo selecionado!")
        return

    try:
        input_pdf = PdfReader(arquivo_pdf)
        total_paginas = len(input_pdf.pages)
        if total_paginas == 0:
            messagebox.showerror("Erro", "O arquivo PDF não contém páginas!")
            return

        selecao = simpledialog.askstring(
            "Seleção de Páginas",
            f"Digite os números das páginas a incluir, separados por vírgulas (ex: 1,2,5-7). Total de páginas: {total_paginas}."
        )

        if not selecao:
            messagebox.showwarning("Aviso", "Nenhuma seleção de páginas feita!")
            return

        paginas_selecionadas = []
        try:
            for parte in selecao.split(","):
                if "-" in parte:
                    inicio, fim = map(int, parte.split("-"))
                    paginas_selecionadas.extend(range(inicio - 1, fim))
                else:
                    paginas_selecionadas.append(int(parte) - 1)

            paginas_selecionadas = sorted(set(paginas_selecionadas))
            if any(p < 0 or p >= total_paginas for p in paginas_selecionadas):
                raise ValueError("Seleção de páginas fora do intervalo válido.")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro na seleção de páginas: {e}")
            return

        writer = PdfWriter()
        for page_num in paginas_selecionadas:
            writer.add_page(input_pdf.pages[page_num])

        pasta_saida = os.path.dirname(arquivo_pdf)
        output_path = os.path.join(pasta_saida, "PDF_Selecionado.pdf")
        with open(output_path, "wb") as output_pdf:
            writer.write(output_pdf)

        messagebox.showinfo("Sucesso", f"PDF gerado com sucesso!\nArquivo salvo em: {output_path}")

        if messagebox.askyesno("Abrir Diretório", "Deseja abrir o diretório onde o arquivo foi salvo?"):
            abrir_diretorio(pasta_saida)

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o arquivo PDF: {e}")

def imagem_para_pdf():
    root = Tk()
    root.withdraw()

    arquivos_imagem = filedialog.askopenfilenames(
        title="Selecione as imagens",
        filetypes=[("Arquivos de Imagem", "*.png;*.jpg;*.jpeg")]
    )

    if not arquivos_imagem:
        messagebox.showwarning("Aviso", "Nenhuma imagem selecionada!")
        return

    try:
        imagens = [Image.open(img).convert('RGB') for img in arquivos_imagem]

        output_pdf = filedialog.asksaveasfilename(
            title="Salvar PDF como",
            defaultextension=".pdf",
            filetypes=[("Arquivo PDF", "*.pdf")]
        )

        if not output_pdf:
            messagebox.showwarning("Aviso", "Nenhum local selecionado para salvar o PDF!")
            return

        imagens[0].save(output_pdf, save_all=True, append_images=imagens[1:])
        messagebox.showinfo("Sucesso", f"PDF criado com sucesso em: {output_pdf}")

        if messagebox.askyesno("Abrir Diretório", "Deseja abrir o diretório onde o arquivo foi salvo?"):
            abrir_diretorio(os.path.dirname(output_pdf))

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter imagens para PDF: {e}")

def pdf_para_imagem():
    import fitz  # PyMuPDF para extrair imagens

    root = Tk()
    root.withdraw()

    arquivo_pdf = filedialog.askopenfilename(
        title="Selecione o arquivo PDF",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )

    if not arquivo_pdf:
        messagebox.showwarning("Aviso", "Nenhum arquivo selecionado!")
        return

    pasta_saida = filedialog.askdirectory(
        title="Selecione o diretório para salvar as imagens"
    )

    if not pasta_saida:
        messagebox.showwarning("Aviso", "Nenhuma pasta de destino selecionada!")
        return

    try:
        pdf_documento = fitz.open(arquivo_pdf)
        for i, pagina in enumerate(pdf_documento):
            imagem = pagina.get_pixmap()
            imagem_path = os.path.join(pasta_saida, f"pagina_{i + 1}.png")
            imagem.save(imagem_path)

        messagebox.showinfo("Sucesso", f"Imagens extraídas com sucesso em: {pasta_saida}")

        if messagebox.askyesno("Abrir Diretório", "Deseja abrir o diretório onde as imagens foram salvas?"):
            abrir_diretorio(pasta_saida)

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter PDF para imagens: {e}")

def menu_principal():
    root = Tk()
    root.title("Conversor de PDF")
    root.geometry("300x200")

    Button(root, text="Dividir PDF", command=dividir_pdf, width=30, height=2).pack(pady=5)
    Button(root, text="Imagens para PDF", command=imagem_para_pdf, width=30, height=2).pack(pady=5)
    Button(root, text="PDF para Imagens", command=pdf_para_imagem, width=30, height=2).pack(pady=5)
    Button(root, text="Sair", command=root.quit, width=30, height=2).pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    menu_principal()
