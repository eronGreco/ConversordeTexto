import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import webbrowser
from threading import Thread
from file_processing import processar_em_lote, processar_arquivo, extrair_textos_sem_formatar, center_window, abrir_configuracao_formatacao_lote, abrir_configuracao_formatacao_individual, mostrar_mensagem_finalizacao, abrir_explorador_textos_formatados
from preview import mostrar_preview

if hasattr(sys, '_MEIPASS'):
    tkdnd_path = os.path.join(sys._MEIPASS, 'tkdnd2.9.4')
else:
    tkdnd_path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'tkdnd2.9.4'))
os.environ['TCLLIBPATH'] = tkdnd_path

def iniciar_interface_grafica():
    global root, arquivos_selecionados, modo_processamento, check_vars, incluir_imagens, selecionar_todos_estado
    root = TkinterDnD.Tk()  # Usar TkinterDnD para drag-and-drop
    root.title("Conversor de Texto")
    center_window(root, 1200, 600)

    label = tk.Label(root, text="Arraste e solte arquivos aqui", width=500, height=10, bg="lightblue")
    label.pack(pady=0)

    arquivos_selecionados = []
    check_vars = []
    modo_processamento = ""
    incluir_imagens = tk.BooleanVar()
    selecionar_todos_estado = False  # Variável para controlar o estado de seleção

    def drop(event):
        arquivos = root.tk.splitlist(event.data)
        for arquivo in arquivos:
            if arquivo not in arquivos_selecionados:
                arquivos_selecionados.append(arquivo)
                check_vars.append(tk.BooleanVar())
        update_lista_arquivos()

    def update_lista_arquivos():
        for widget in inner_frame.winfo_children():
            widget.destroy()

        for idx, arquivo in enumerate(arquivos_selecionados):
            frame = tk.Frame(inner_frame)
            frame.pack(fill=tk.X)

            btn_delete = tk.Button(frame, text="🗑", command=lambda f=arquivo: excluir_arquivo(f))
            btn_delete.pack(side=tk.LEFT)

            btn_preview = tk.Button(frame, text="👁", command=lambda f=arquivo: mostrar_preview(root, f))
            btn_preview.pack(side=tk.LEFT)

            checkbutton = tk.Checkbutton(frame, variable=check_vars[idx])
            checkbutton.pack(side=tk.LEFT)

            nome_arquivo = os.path.basename(arquivo)
            label = tk.Label(frame, text=nome_arquivo, anchor='w')
            label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        lbl_quantidade_arquivos.config(text=f"Quantidade de arquivos: {len(arquivos_selecionados)}")

    def excluir_arquivo(arquivo):
        idx = arquivos_selecionados.index(arquivo)
        arquivos_selecionados.pop(idx)
        check_vars.pop(idx)
        update_lista_arquivos()

    def selecionar_arquivos():
        arquivos = filedialog.askopenfilenames(title="Selecione os arquivos para processar",
                                               filetypes=[("Todos os arquivos", "*.*"), ("Arquivos PDF", "*.pdf"),
                                                          ("Documentos do Word", "*.docx"),
                                                          ("Apresentações do PowerPoint", "*.pptx"),
                                                          ("Imagens", "*.jpg;*.jpeg;*.png;*.gif;*.bmp"),
                                                          ("Arquivos de Texto", "*.txt")])
        for arquivo in arquivos:
            if arquivo not in arquivos_selecionados:
                arquivos_selecionados.append(arquivo)
                check_vars.append(tk.BooleanVar())
        update_lista_arquivos()

    def selecionar_todos():
        global selecionar_todos_estado
        selecionar_todos_estado = not selecionar_todos_estado
        for var in check_vars:
            var.set(selecionar_todos_estado)

    def processar_arquivos_lote():
        global modo_processamento
        modo_processamento = "lote"
        if not arquivos_selecionados:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado.")
            return

        incluir_imagens = tk.BooleanVar()

        def iniciar_processamento(overlap, marker):
            output_dir = os.path.join(os.path.dirname(arquivos_selecionados[0]), "Textos formatados")
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            def on_complete():
                abrir_explorador_textos_formatados(output_dir)
                mostrar_mensagem_finalizacao(root)

            thread = Thread(target=processar_em_lote,
                            args=(arquivos_selecionados, output_dir, marker, overlap, incluir_imagens.get()),
                            daemon=True)
            thread.start()
            thread.join(on_complete())

        # Verifica se há arquivos de imagem na lista
        tem_imagens = any(
            arquivo.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp')) for arquivo in arquivos_selecionados)

        abrir_configuracao_formatacao_lote(iniciar_processamento, extrair_textos_sem_formatar_callback,
                                           incluir_imagens if tem_imagens else None)

    def processar_arquivos_individual():
        global modo_processamento, current_index
        modo_processamento = "individual"
        selected_files = [arquivos_selecionados[idx] for idx, var in enumerate(check_vars) if var.get()]
        if not selected_files:
            messagebox.showwarning("Aviso", "Selecione um ou mais arquivos para formatação individual.")
            return

        current_index = 0  # Inicializa o índice antes de iniciar o processo

        output_dir = os.path.join(os.path.dirname(selected_files[0]), "Textos formatados")
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        def processar_proximo_arquivo():
            global current_index
            if current_index < len(selected_files):
                arquivo = selected_files[current_index]

                def iniciar_processamento_individual(overlap, marker):
                    global current_index
                    processar_arquivo(arquivo, output_dir, marker, overlap)
                    current_index += 1
                    processar_proximo_arquivo()

                abrir_configuracao_formatacao_individual(iniciar_processamento_individual,
                                                         extrair_textos_sem_formatar_callback, selected_files)
            else:
                abrir_explorador_textos_formatados(output_dir)
                mostrar_mensagem_finalizacao(root)

        processar_proximo_arquivo()

    def extrair_textos_sem_formatar_callback():
        if modo_processamento == "lote":
            if not arquivos_selecionados:
                messagebox.showwarning("Aviso", "Nenhum arquivo selecionado.")
                return
            output_dir = os.path.join(os.path.dirname(arquivos_selecionados[0]), "Textos extraídos")
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            def on_complete():
                abrir_explorador_textos_formatados(output_dir)
                mostrar_mensagem_finalizacao(root)

            thread = Thread(target=extrair_textos_sem_formatar, args=(arquivos_selecionados, output_dir), daemon=True)
            thread.start()
            thread.join(on_complete())
        elif modo_processamento == "individual":
            selected_files = [arquivos_selecionados[idx] for idx, var in enumerate(check_vars) if var.get()]
            if not selected_files:
                messagebox.showwarning("Aviso", "Nenhum arquivo selecionado.")
                return
            output_dir = os.path.join(os.path.dirname(selected_files[0]), "Textos extraídos")
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            def on_complete():
                abrir_explorador_textos_formatados(output_dir)
                mostrar_mensagem_finalizacao(root)

            thread = Thread(target=extrair_textos_sem_formatar, args=(selected_files, output_dir), daemon=True)
            thread.start()
            thread.join(on_complete())

    def limpar_lista():
        arquivos_selecionados.clear()
        check_vars.clear()
        update_lista_arquivos()

    def abrir_lyriccraft_web():
        # Abrir o arquivo index.html na mesma pasta do executável
        current_dir = os.path.dirname(os.path.abspath(__file__))
        index_path = os.path.join(current_dir, "index.html")
        if os.path.exists(index_path):
            webbrowser.open_new_tab(index_path)
        else:
            messagebox.showerror("Erro", "O arquivo index.html não foi encontrado.")

    # Linha preta acima da lista de arquivos
    tk.Frame(root, height=4, bg="black").pack(fill=tk.X)

    #
    # SCROLL VERTICAL E SCROLL HORIZONTAL
    #

    listbox_frame = tk.Frame(root)
    listbox_frame.pack(pady=10, fill=tk.BOTH, expand=True)

    # Adiciona ScrollBar vertical e horizontal
    canvas = tk.Canvas(listbox_frame, yscrollcommand=lambda *args: scrollbar_y.set(*args),
                       height=200)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar_y = tk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

    inner_frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=inner_frame, anchor='nw')

    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    inner_frame.bind("<Configure>", on_frame_configure)

    # Criação da barra de rolagem horizontal no frame listbox_frame
    #scrollbar_x = tk.Scrollbar(listbox_frame, orient=tk.HORIZONTAL, command=canvas.xview)
    #scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

    # Linha preta abaixo da lista de arquivos
    tk.Frame(root, height=4, bg="black").pack(fill=tk.X)

    lbl_quantidade_arquivos = tk.Label(root, text="Quantidade de arquivos: 0")
    lbl_quantidade_arquivos.pack()

    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # FIM SCROLL HORIZONTAL E VERTICAL



    frame_botoes = tk.Frame(root)
    frame_botoes.pack(pady=10)

    botao_importar = tk.Button(frame_botoes, text="Importar Arquivos", command=selecionar_arquivos)
    botao_limpar = tk.Button(frame_botoes, text="Limpar Lista", command=limpar_lista)
    botao_selecionar_todos = tk.Button(frame_botoes, text="Selecionar Todos", command=selecionar_todos)
    botao_processar_lote = tk.Button(frame_botoes, text="Processar em Lote", command=processar_arquivos_lote)
    botao_lyriccraft_web = tk.Button(frame_botoes, text="LyricCraft Web", command=abrir_lyriccraft_web)
    botao_processar_individual = tk.Button(frame_botoes, text="Processar Individualmente", command=processar_arquivos_individual)

    botao_importar.grid(row=0, column=0, padx=5, pady=5)
    botao_limpar.grid(row=0, column=2, padx=5, pady=5)
    botao_selecionar_todos.grid(row=0, column=1, padx=5, pady=5)
    botao_processar_lote.grid(row=0, column=3, padx=5, pady=5)
    # botao_lyriccraft_web.grid(row=0, column=5, padx=5, pady=5)
    botao_processar_individual.grid(row=0, column=4, padx=5, pady=5)

    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', drop)

    root.mainloop()

if __name__ == "__main__":
    iniciar_interface_grafica()