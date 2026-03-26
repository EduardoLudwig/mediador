import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import sys
import os
import pythoncom
from robo_mediador_original import main, stop_execution, load_last_work_path, save_last_work_path, load_last_excel_path, save_last_excel_path

# adicionar o icone de giro para mostrar que esta rodando (isso já fiz no rodo dos ccts do comuniq, só copiar e colar)


# mensagem de alerta para caso de falta de espaço (isso seria para a rede)

# caso ler fechada ou #N/D em QUALQUER UM DOS ITENS DEVE PULAR AQUELA LOJA  

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


verde_qq = "#277231"
verde_qq_hover = "#3E8E49"

ctk.set_appearance_mode("dark")

class App:
    def __init__(self, root, auto_start=False):
        self.root = root
        self.root.title("Robô Mediador")
        self.root.geometry("800x450")
        self.root.resizable(False, False)
        self.root.configure(fg_color=["gray16", "gray16"]) # primeiro valor é para light mode e segundo é para dark mode
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
        self.top_frame = ctk.CTkFrame(self.root, fg_color="transparent", height=40)
        self.top_frame.pack(side="top", fill="x", padx=10, pady=5)

        self.top_frame.grid_columnconfigure(0, weight=1)
        self.top_frame.grid_columnconfigure(1, weight=0) 

        self.label_download_path = ctk.CTkLabel(self.top_frame, text="Nenhuma pasta de trabalho selecionada.", font=("arial", 12, "bold"))
        self.label_download_path.grid(row=0, column=0, sticky="n", padx=(60, 10), pady=10)

        self.botao_config = ctk.CTkButton(
            self.top_frame, 
            text="⚙️",
            command=self.selecionar_pasta_downloads,
            font=ctk.CTkFont(size=24),
            fg_color="transparent",
            hover_color=verde_qq_hover,
            width=30,
            height=30
        )
        
        self.botao_config.grid(row=0, column=1, sticky="e")

        self.frame_fundo = ctk.CTkFrame(self.root, fg_color="transparent")
        self.frame_fundo.pack(fill="both", expand=True, padx=20, pady=20)

        self.loja_inicial_frame = ctk.CTkFrame(self.frame_fundo, fg_color="transparent")

        self.loja_inicial_frame.pack(side="bottom", pady=(10, 0))

        self.central_frame = ctk.CTkFrame(self.frame_fundo, fg_color="transparent")
        self.central_frame.place(relx=0.5, rely=0.5, anchor="center")  

        self.botao_frame = ctk.CTkFrame(self.central_frame, fg_color="transparent")
        self.botao_frame.pack(pady=10)

        self.label_frame = ctk.CTkFrame(self.frame_fundo, fg_color="transparent")
        self.label_frame.place(relx=0.5, rely=0.050, anchor="n")  

        self.label = ctk.CTkLabel(
            self.label_frame,
            text="SELECIONE A BASE DE DADOS PARA PROCESSAMENTO",
            text_color=["#F3F5E8", "#F0F8ED"],
            fg_color=verde_qq,
            corner_radius=6,
            font=("Inter", 16, "bold")
        )
        self.label.pack(pady=(0, 20))

        self.botao_excel = ctk.CTkButton(
            self.botao_frame,
            text="SELECIONAR BASE DE DADOS",
            command=self.selecionar_excel,
            fg_color=verde_qq,
            hover_color=verde_qq_hover,
            font=("Inter", 12, "bold")
        )
        self.botao_excel.pack(side="left", padx=30)

        self.botao_iniciar = ctk.CTkButton(
            self.botao_frame,
            text="INICIAR PROCESSAMENTO",
            state="disabled",
            command=self.iniciar_processamento,
            fg_color=verde_qq,
            hover_color=verde_qq_hover,
            font=("Inter", 12, "bold")
        )
        self.botao_iniciar.pack(side="right", padx=30)
        
        self.loja_inicial_frame = ctk.CTkFrame(self.loja_inicial_frame, fg_color="transparent")
        self.loja_inicial_frame.pack(pady=(40, 10))

        self.label_loja_inicial = ctk.CTkLabel(self.loja_inicial_frame, text="Iniciar a partir da loja:", font=("Inter", 18, "bold"))
        self.label_loja_inicial.pack(side="left", padx=5)

        self.entry_loja_inicial = ctk.CTkEntry(
            self.loja_inicial_frame,
            width=200,
            placeholder_text="Padrão: Início da lista (Loja 1)"
        )
        self.entry_loja_inicial.pack(side="left")


        self.status = ctk.CTkLabel(self.central_frame, text="", fg_color="transparent", corner_radius=6, font=("Inter", 12, "bold"))

        self.status.pack(pady=10)

        self.processing_store_label = ctk.CTkLabel(self.central_frame, text="")
        self.processing_store_label.pack(pady=10)

        self.caminho_arquivo = None
        self.caminho_download = None
        self.process_running = False 
        self.browser_instance = None

        self.load_initial_settings()

        if auto_start:
            self.auto_start_process()

    def auto_start_process(self):
        last_excel_path = load_last_excel_path()

        if not self.caminho_download:
            messagebox.showwarning("Atenção", "Pasta de trabalho não definida. Não é possível iniciar automaticamente.")
            return

        if not last_excel_path or not os.path.exists(last_excel_path):
            messagebox.showwarning("Atenção", "Caminho do último Excel não encontrado ou inválido. Não é possível iniciar automaticamente.")
            return
        
        self.caminho_arquivo = last_excel_path
        self.iniciar_processamento()
    
    def load_last_excel_path_from_config(self):
        from robo_mediador import load_last_excel_path
        return load_last_excel_path()


    def load_initial_settings(self):
        last_path = load_last_work_path()
        if last_path and os.path.exists(last_path):
            self.caminho_download = last_path
            print(f"Pasta de trabalho anterior carregada: {last_path}")

            # Atualiza o texto do label para mostrar a pasta carregada
            caminho_exibicao = "..." + last_path[-50:] if len(last_path) > 50 else last_path
            self.label_download_path.configure(text=f"Pasta de Trabalho: {caminho_exibicao}")
        else:
            print("Nenhuma pasta de trabalho anterior encontrada ou o caminho é inválido.")

        last_excel_path = load_last_excel_path()
        if last_excel_path and os.path.exists(last_excel_path):
            self.caminho_arquivo = last_excel_path
            print(f"Caminho do Excel anterior carregado: {last_excel_path}")

            if self.caminho_download:
                self.botao_iniciar.configure(state="normal")
                self.status.configure(text="Caminho do Excel carregado.")
        else:
            print("Nenhum caminho do Excel anterior encontrado.")
                

    def selecionar_excel(self):
        self.caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xlsb")])
        if self.caminho_arquivo:
            save_last_excel_path(self.caminho_arquivo)
            self.status.configure(text="Arquivo selecionado com sucesso.")
            if self.caminho_download:
                self.botao_iniciar.configure(state="normal")
        else:
            self.status.configure(text="Arquivo deve ser em formato excel (.xlsx ou .xlsb).") 
            self.botao_iniciar.configure(state="disabled")

    def iniciar_processamento(self):

        if not self.caminho_arquivo or not self.caminho_download:
            messagebox.showwarning("Atenção", "Por favor, selecione a base de dados Excel e a pasta de trabalho antes de iniciar.")
            return

        loja_inicial = self.entry_loja_inicial.get()

        self.status.configure(text="Executando processamento, por favor aguarde.", font=("Inter", 12, "bold"))
        self.botao_iniciar.configure(state="disabled")
        self.botao_excel.configure(state="disabled")
        self.entry_loja_inicial.configure(state="disabled")
        self.process_running = True

        self.thread = threading.Thread(target=self.executar_main, args=(self.caminho_download, loja_inicial, self.update_processing_store_label,), daemon=True)
        self.thread.start()

    def executar_main(self, download_path, loja_inicial, update_callback):
        pythoncom.CoInitialize()
        try:
            main(self.caminho_arquivo, base_path_param=download_path, loja_inicial=loja_inicial, update_callback=update_callback) 
            self.status.configure(text="Processo concluído com sucesso.")
            self.processing_store_label.configure(text="")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
            self.status.configure(text="Erro durante a execução.")
            self.processing_store_label.configure(text="")
        finally:
            self.process_running = False 
            self.botao_iniciar.configure(state="normal")
            self.botao_excel.configure(state="normal")
            self.entry_loja_inicial.configure(state="normal")
            if self.browser_instance:
                try:
                    self.browser_instance.quit()
                except Exception as e:
                    print(f"Erro ao fechar o navegador no finally da interface: {e}")

            pythoncom.CoUninitialize()

    def update_processing_store_label(self, loja_info):
        self.processing_store_label.configure(text=f"Loja: {loja_info}")
    
    
    def selecionar_pasta_downloads(self):
        self.caminho_download = filedialog.askdirectory(title="Selecione a pasta de trabalho principal")
        if self.caminho_download:
            save_last_work_path(self.caminho_download)

            caminho_exibicao = "..." + self.caminho_download[-50:] if len(self.caminho_download) > 50 else self.caminho_download
            self.label_download_path.configure(text=f"Pasta selecionada: {caminho_exibicao}")
            
            if self.caminho_arquivo:
                self.botao_iniciar.configure(state="normal")
        else:
            self.label_download_path.configure(text="Nenhuma pasta de trabalho selecionada.")
            self.botao_iniciar.configure(state="disabled")


    def on_close(self):
        if self.process_running:
            if messagebox.askyesno("Confirmação", "O processo ainda não foi concluido. Deseja sair e encerrar a aplicação?"):
                stop_execution()
                self.root.destroy()
                sys.exit()
            else:
                return
        else:
            self.root.destroy()
            sys.exit()

if __name__ == "__main__":
    root = ctk.CTk()
    should_auto_start = '--auto-start' in sys.argv
    app = App(root, auto_start=should_auto_start)
    root.mainloop()
