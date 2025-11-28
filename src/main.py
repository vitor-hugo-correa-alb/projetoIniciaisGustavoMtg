import customtkinter as ctk
import os
import sys
import logging
from logging.handlers import RotatingFileHandler
from tkinter import filedialog, Tk, messagebox

# Determina o diretório base a ser usado para buscar templates e criar logs.
# Requisito: quando executável (.exe) rodando, usar o diretório do executável.
def get_base_dir_for_logging():
    if getattr(sys, "frozen", False):
        # Em um exe, usamos o diretório do executável (por exemplo C:\path\to\dist)
        return os.path.dirname(sys.executable)
    # Em desenvolvimento (executando como módulo), usamos a raiz do projeto (pai de src)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Substitui o token {{BARRA}} por '/' para exibição (não altera nomes de arquivo)
def replace_bar_placeholder(text: str) -> str:
    if text is None:
        return text
    return text.replace("{{BARRA}}", "/")

# Configura logging antes de importar modules que possam logar
BASE_DIR = get_base_dir_for_logging()
LOGS_DIR = os.path.join(BASE_DIR, "logs")
os.makedirs(LOGS_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOGS_DIR, "logs.txt")

formatter = logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s", "%Y-%m-%d %H:%M:%S")
root_logger = logging.getLogger()
root_logger.setLevel(logging.INFO)

file_handler = RotatingFileHandler(LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8")
file_handler.setFormatter(formatter)
root_logger.addHandler(file_handler)

console_handler = logging.StreamHandler()
console_handler.setFormatter(formatter)
root_logger.addHandler(console_handler)

logger = logging.getLogger(__name__)
logger.info("Inicializando aplicação — base_dir=%s", BASE_DIR)

# Agora importamos generate_word (que usa logging)
from .generate_word import gerar_documento, salvar_documento

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

def get_base_dir():
    """
    Retorna o diretório base onde procurar recursos (templates).
    Quando executado como exe (PyInstaller --onefile), usa o diretório do executável.
    Quando executado em desenvolvimento (python -m src.main), usa a raiz do projeto.
    """
    return BASE_DIR

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Gerador Inicial - Automação Jurídica")
        self.geometry("1100x700")
        self.minsize(800, 500)

        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=12, pady=12)

        # Layout horizontal para os dois painéis lado a lado (mesmo tamanho visual)
        h_frame = ctk.CTkFrame(main_frame)
        h_frame.pack(fill="both", expand=True)

        # ---- PREÂMBULO DA AÇÃO COM SCROLL ----
        preambulo_frame = ctk.CTkFrame(h_frame, corner_radius=8)
        preambulo_frame.pack(side="left", fill="both", expand=True, padx=(0, 6), pady=0)
        preambulo_frame.pack_propagate(False)

        preambulo_label = ctk.CTkLabel(preambulo_frame, text="Preâmbulo da Ação", font=("Arial", 18, "bold"))
        preambulo_label.pack(pady=(12, 6))

        # Adicione um CTkScrollableFrame para os campos
        preambulo_scroll = ctk.CTkScrollableFrame(preambulo_frame)
        preambulo_scroll.pack(fill="both", expand=True, padx=8, pady=(0,8))

        self.campos = {}
        nomes_campos = [
            "Nome reclamante", "Profissao / Cargo", "Data de nascimento", "Nome da mae",
            "Número do pis", "Número da ctps", "Número rg", "Número do cpf", "Rua do reclamante",
            "Número da casa do reclamante e complemento", "Bairro reclamante", "Cep reclamante",
            "Nome da reclamada", "Empresa processada", "Numero de cnpj da reclamada",
            "Endereço reclamada", "Complemento reclamada", "Bairro reclamada", "Cep reclamada"
        ]
        for nome in nomes_campos:
            ctk.CTkLabel(preambulo_scroll, text=nome).pack(anchor="w", padx=6, pady=(6,0))
            entry = ctk.CTkEntry(preambulo_scroll)
            entry.pack(anchor="w", fill="x", pady=(0,8), padx=6)
            self.campos[nome] = entry

        # ---- MODELOS DE PEDIDO COM SCROLL ----
        modelos_frame = ctk.CTkFrame(h_frame, corner_radius=8)
        modelos_frame.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=0)
        modelos_frame.pack_propagate(False)

        modelos_label = ctk.CTkLabel(modelos_frame, text="Modelos de Pedido", font=("Arial", 18, "bold"))
        modelos_label.pack(pady=(12, 6))

        # guardamos o scroll como atributo para poder recriar os checkboxes quando atualizar
        self.modelos_scroll = ctk.CTkScrollableFrame(modelos_frame)
        self.modelos_scroll.pack(fill="both", expand=True, padx=8, pady=(0,8))

        # estrutura que armazenará os modelos: lista de tuples (nome_exibicao_raw, var_boolean, caminho)
        # nome_exibicao_raw = filename without extension, may contain {{BARRA}} token
        self.modelos_vars = []
        self.modelos_widgets = []  # referências aos widgets de checkbox (para poder remover ao recarregar)
        self.selecionados_ordem = []  # mantém a ordem em que o usuário marcou (se usar)

        # Carrega modelos iniciais
        self.carregar_modelos()

        # ---- FRAME DE BOTOES (fora dos painéis), centralizados ----
        botoes_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        botoes_frame.pack(fill="x", pady=(12, 0))

        buttons_container = ctk.CTkFrame(botoes_frame, fg_color="transparent")
        buttons_container.pack()

        self.botao_gerar = ctk.CTkButton(buttons_container, text="Gerar Inicial", width=180, height=46, font=("Arial", 14), command=self.gerar_inicial)
        self.botao_gerar.pack(side="left", padx=(0, 14))

        self.botao_limpar = ctk.CTkButton(buttons_container, text="Limpar", width=140, height=46, font=("Arial", 14), fg_color="#d9534f", hover_color="#c9302c", command=self.limpar_campos)
        self.botao_limpar.pack(side="left")

        # botão para atualizar a lista de modelos (útil durante testes)
        self.botao_atualizar = ctk.CTkButton(buttons_container, text="Atualizar Modelos", width=160, height=36, font=("Arial",12), command=self.carregar_modelos)
        self.botao_atualizar.pack(side="left", padx=(12,0))

    def carregar_modelos(self):
        """
        (Re)carrega a lista de .docx em templates/modelos e recria os checkboxes no scroll.
        Preserva seleções/ordem anteriores quando possível.
        """
        prev_selected = {c: v.get() for (_, v, c) in self.modelos_vars}
        prev_order = list(self.selecionados_ordem)

        # limpa estruturas anteriores (widgets)
        for w in self.modelos_widgets:
            try:
                w.destroy()
            except Exception:
                pass
        self.modelos_widgets = []
        self.modelos_vars = []
        self.selecionados_ordem = []

        base_dir = get_base_dir()
        modelos_dir = os.path.join(base_dir, "templates", "modelos")
        if not os.path.isdir(modelos_dir):
            os.makedirs(modelos_dir)

        arquivos = [f for f in os.listdir(modelos_dir) if f.lower().endswith('.docx')]
        arquivos.sort()  # ordenação estável por nome (você pode escolher outra)
        for nome_arquivo in arquivos:
            caminho_arquivo = os.path.join(modelos_dir, nome_arquivo)
            var = ctk.BooleanVar(value=prev_selected.get(caminho_arquivo, False))
            nome_exibicao_raw = os.path.splitext(nome_arquivo)[0]  # raw name, may include {{BARRA}}
            display_name = replace_bar_placeholder(nome_exibicao_raw)
            chk = ctk.CTkCheckBox(
                self.modelos_scroll,
                text=display_name,
                variable=var,
                command=lambda n=nome_exibicao_raw, p=caminho_arquivo, v=var: self.atualizar_ordem_selecao(n, p, v)
            )
            chk.pack(anchor="w", pady=(0,8), padx=6)
            self.modelos_widgets.append(chk)
            # store the raw name so we can replace token when needed (and keep consistency)
            self.modelos_vars.append((nome_exibicao_raw, var, caminho_arquivo))

        # rebuild selecionados_ordem preserving prev_order sequence first
        for nome, caminho in prev_order:
            for (n, v, c) in self.modelos_vars:
                if c == caminho and v.get():
                    if (n, c) not in self.selecionados_ordem:
                        self.selecionados_ordem.append((n, c))
        # then append any other currently selected that weren't in prev_order
        for (n, v, c) in self.modelos_vars:
            if v.get() and (n, c) not in self.selecionados_ordem:
                self.selecionados_ordem.append((n, c))

        logger.info("[main.carregar_modelos] modelos carregados: %s", [(n,c) for (n,_,c) in self.modelos_vars])
        logger.debug("[main.carregar_modelos] selecionados_ordem reconstruido: %s", self.selecionados_ordem)

        # Atualiza display das labels dos modelos para mostrar ordem quando marcados
        self.update_modelos_display()

    def atualizar_ordem_selecao(self, nome, caminho, var):
        valor = var.get()
        if valor:
            if (nome, caminho) not in self.selecionados_ordem:
                self.selecionados_ordem.append((nome, caminho))
        else:
            self.selecionados_ordem = [(n, c) for n, c in self.selecionados_ordem if c != caminho]

        # Atualiza visualmente os textos dos checkboxes para mostrar a ordem atual
        self.update_modelos_display()

    def update_modelos_display(self):
        """
        Atualiza o texto de cada checkbox para exibir a ordem de geração quando selecionado.
        Ex.: "1 - Nome Modelo", "2 - Outro Modelo", ou apenas "Nome Modelo" quando não selecionado.
        Usa o caminho do arquivo para comparar e encontrar a posição correta na lista self.selecionados_ordem.
        """
        # cria mapa caminho -> posição (1-based)
        ordem_map = {}
        for idx, (n, c) in enumerate(self.selecionados_ordem, start=1):
            ordem_map[c] = idx

        # atualiza cada checkbox na mesma ordem de self.modelos_vars / self.modelos_widgets
        for (nome_exibicao_raw, var, caminho), chk in zip(self.modelos_vars, self.modelos_widgets):
            try:
                display_name = replace_bar_placeholder(nome_exibicao_raw)
                if var.get() and caminho in ordem_map:
                    pos = ordem_map[caminho]
                    chk.configure(text=f"{pos} - {display_name}")
                else:
                    chk.configure(text=display_name)
            except Exception:
                # não deixar que erros na UI quebrem a lista
                try:
                    chk.configure(text=replace_bar_placeholder(nome_exibicao_raw))
                except Exception:
                    pass

    def limpar_campos(self):
        # Limpa entradas de texto
        for nome, entry in self.campos.items():
            try:
                entry.delete(0, 'end')
            except Exception:
                entry.set("") if hasattr(entry, "set") else None

        # Desmarca todos os checkboxes e limpa a ordem de seleção
        for nome, var, caminho in self.modelos_vars:
            try:
                var.set(False)
            except Exception:
                pass
        self.selecionados_ordem = []
        # atualizar display
        self.update_modelos_display()

        messagebox.showinfo("Limpar", "Campos e seleções foram limpos.")
        logger.info("Campos limpos e seleções removidas pelo usuário.")

    def gerar_inicial(self):
        # NÃO recarregar modelos aqui (isso apagava as seleções); se quiser atualizar manualmente use o botão "Atualizar Modelos"

        # Recolhe dados do preâmbulo
        dados = {campo: entry.get() for campo, entry in self.campos.items()}

        # Primeiro tenta usar a ordem interna (se o usuário estava usando isso)
        pedidos_ordenados = list(self.selecionados_ordem)  # copia para não alterar original

        # Se estiver vazia, reconstrói a lista a partir dos checkboxes marcados (sem ordem)
        if not pedidos_ordenados:
            pedidos_ordenados = [(nome, caminho) for (nome, var, caminho) in self.modelos_vars if var.get()]

        logger.info("[main.gerar_inicial] pedidos_ordenados: %s", pedidos_ordenados)
        logger.debug("[main.gerar_inicial] modelos detectados: %s", [(nome, caminho) for (nome, _, caminho) in self.modelos_vars])

        if not pedidos_ordenados:
            if not messagebox.askyesno("Nenhum modelo selecionado", "Nenhum modelo selecionado. Deseja gerar apenas a base (modelo_base)?"):
                logger.info("Usuário cancelou geração sem modelos selecionados.")
                return

        base_dir = get_base_dir()
        caminho_template = os.path.join(base_dir, 'templates', 'modelo_base.docx')
        if not os.path.exists(caminho_template):
            logger.error("Template não encontrado em %s", caminho_template)
            messagebox.showerror("Erro", f"Template não encontrado em {caminho_template}")
            return

        try:
            doc = gerar_documento(caminho_template, dados, pedidos_ordenados, 6)  # numeracao padrão 6
        except Exception as e:
            logger.exception("Erro ao gerar documento")
            messagebox.showerror("Erro na geração", f"Ocorreu um erro ao gerar o documento:\n{e}")
            return

        root = Tk()
        root.withdraw()
        caminho_destino = filedialog.asksaveasfilename(
            title="Salvar documento gerado",
            filetypes=[("Word Document", "*.docx")],
            defaultextension=".docx"
        )
        root.destroy()
        if caminho_destino:
            try:
                salvar_documento(doc, caminho_destino)
                logger.info("Documento salvo em %s", caminho_destino)
                messagebox.showinfo("Sucesso", f"Documento salvo em:\n{caminho_destino}")
            except Exception as e:
                logger.exception("Erro ao salvar documento")
                messagebox.showerror("Erro ao salvar", f"Não foi possível salvar o arquivo:\n{e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()