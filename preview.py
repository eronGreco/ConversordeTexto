import os
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
from docx import Document
import fitz  # PyMuPDF

def mostrar_preview(root, arquivo):
    if not os.path.isfile(arquivo):
        return

    preview_window = tk.Toplevel(root)
    preview_window.title("Pré-visualização")
    preview_window.geometry("600x800")  # Ajustei o tamanho da janela para uma visualização melhor
    center_window(preview_window, 600, 800)

    # Frame para a área de visualização com barras de rolagem
    frame = ttk.Frame(preview_window)
    frame.pack(fill=tk.BOTH, expand=1)

    # Canvas para a imagem ou texto
    canvas = tk.Canvas(frame)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

    # Barra de rolagem vertical
    v_scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
    v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # Barra de rolagem horizontal
    h_scrollbar = ttk.Scrollbar(preview_window, orient=tk.HORIZONTAL, command=canvas.xview)
    h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

    canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
    canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    # Frame interno para o canvas
    inner_frame = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=inner_frame, anchor="nw")

    def load_image(im):
        img = Image.open(im)
        return ImageTk.PhotoImage(img)

    def load_pdf(pdf_path):
        pdf = fitz.open(pdf_path)
        page = pdf.load_page(0)
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return ImageTk.PhotoImage(img)

    try:
        if arquivo.lower().endswith('.pdf'):
            img = load_pdf(arquivo)
            panel = tk.Label(inner_frame, image=img)
            panel.image = img
            panel.pack()
        elif arquivo.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp')):
            img = load_image(arquivo)
            panel = tk.Label(inner_frame, image=img)
            panel.image = img
            panel.pack()
        elif arquivo.lower().endswith('.docx'):
            doc = Document(arquivo)
            texto = "\n".join([para.text for para in doc.paragraphs])
            lbl = tk.Label(inner_frame, text=texto, wraplength=600)
            lbl.pack()
        else:
            with open(arquivo, 'r', encoding='iso-8859-1') as f:
                texto = f.read()
            lbl = tk.Label(inner_frame, text=texto, wraplength=600)
            lbl.pack()
    except Exception as e:
        print(f"Erro ao mostrar pré-visualização: {e}")

    # Garantir que a barra de rolagem horizontal esteja na extremidade inferior da janela
    h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    position_top = int(screen_height / 2 - height / 2)
    position_right = int(screen_width / 2 - width / 2)
    window.geometry(f'{width}x{height}+{position_right}+{position_top}')

# Exemplo de uso
if __name__ == "__main__":
    root = tk.Tk()
    root.title("App de Pré-visualização")
    root.geometry("300x200")

    btn = tk.Button(root, text="Mostrar Pré-visualização", command=lambda: mostrar_preview(root, "seuarquivo.pdf"))
    btn.pack(pady=20)

    root.mainloop()
