import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, scrolledtext
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from ttkbootstrap.style import Style    
from urllib.parse import quote
from time import sleep
from datetime import datetime
import pandas as pd
import webbrowser
import threading
import pyautogui
import json
import os
import logging
import openpyxl
import tkinter.messagebox
import re

logging.basicConfig(filename='mesagens_enviadas.log', level=logging.INFO, format='%(asctime)s - %(message)s')

def limpar_telefone(valor):
    if not valor:
        return None
    apenas_numeros = re.sub(r'\D', '', str(valor))
    try:
        return int(apenas_numeros)
    except ValueError:
        return None

def add_placeholder(entry, placeholder_text, color='grey'):
    entry.insert(0, placeholder_text)
    entry.configure(foreground=color)

    def on_focus_in(event):
        if entry.get() == placeholder_text:
            entry.delete(0, "end")
            entry.configure(foreground='')

    def on_focus_out(event):
        if entry.get() == "":
            entry.insert(0, placeholder_text)
            entry.configure(foreground=color)

    entry.bind("<FocusIn>", on_focus_in)
    entry.bind("<FocusOut>", on_focus_out)

class ToolTip(object):
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text):
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 25
        y = y + cy + self.widget.winfo_rooty() + 25
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                       background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                       font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

def create_tooltip(widget, text):
    toolTip = ToolTip(widget, text)
    def enter(event):
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)

def carregar_cursos():
    if os.path.exists("config_cursos.json"):
        with open("config_cursos.json", "r", encoding='utf-8') as f:
            return json.load(f)
    else:
        modelo_padrao = {
            "Geral": ["Exemplo Curso 1", "Exemplo Curso 2"]
        }
        with open("config_cursos.json", "w", encoding='utf-8') as f:
            json.dump(modelo_padrao, f, indent=4, ensure_ascii=False)
        return modelo_padrao

def salvar_cursos_json(dados):
    with open("config_cursos.json", "w", encoding='utf-8') as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)

def carregar_mensagem_padrao():
    arquivo = "mensagem_padrao.txt"
    modelo_padrao = (
        "Olá *{nome}.* Nós somos da AMTECH - Agência Maringá de Tecnologia e Inovação. "
        "entramos em contato porque você demonstrou interesse em cursos de tecnologia preenchendo um formulário.📋\n\n"
        "Nós iremos iniciar em parceria com o *{parceiro}*, o curso:\n\n"
        "🌟*{curso}*.🌟\n\n"
        "Todos podem participar desde que sejam maior de *{idade_minima}* anos e tenham a escolaridade mínima do Ensino Fundamental Completo.🎓\n\n"
        "🎯 Duração do curso: *{duracao}*\n\n"
        "🕒 Horário: *{horario}*\n\n"
        "⚠️ Atenção: As vagas são limitadas! Responda o mais rápido possível! 🏃‍♂️💨 📢*\n\n"
        "*📍Local: Acesso 1 | Piso Superior Terminal Urbano - Av. Tamandaré, 600 - Zona 01, Maringá🗺️ -*\n\n"
        "*🏫 MODALIDADE: curso é PRESENCIAL E 100% GRATUITO! 🎉*\n\n"
        "Qualquer dúvida, estamos à disposição! Esperamos você! 😉"
    )
    
    if os.path.exists(arquivo):
        with open(arquivo, "r", encoding='utf-8') as f:
            return f.read()
    else:
        with open(arquivo, "w", encoding='utf-8') as f:
            f.write(modelo_padrao)
        return modelo_padrao

def salvar_mensagem_padrao(texto):
    with open("mensagem_padrao.txt", "w", encoding='utf-8') as f:
        f.write(texto)

def load_last_line():
    if os.path.exists("last_line.json"):
        with open("last_line.json", "r") as f:
            data = json.load(f)
            return data.get("Ultima_linha_enviada", 0)
    return 0

def save_last_line(last_line):
    with open("last_line.json", "w") as f:
        json.dump({"Ultima_linha_enviada": last_line}, f)

def load_settings():
    if os.path.exists("settings.json"):
        with open("settings.json", "r") as f:
            return json.load(f)
    else:
        return {"theme": "journal"} 
    
def save_settings(settings):
    with open("settings.json", "w") as f:
        json.dump(settings, f)

def carregar_numeros_enviados():
    if os.path.exists("numeros_enviados.json"):
        with open("numeros_enviados.json", "r") as f:
            return set(json.load(f))
    return set()

def salvar_numeros_enviados(numeros):
    with open("numeros_enviados.json", "w") as f:
        json.dump(list(numeros), f)

class MessageEditor(ttkb.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Editor de Mensagem Padrão")
        self.geometry("700x600")
        
        ttkb.Label(self, text="Edite o modelo da mensagem abaixo.", font=("Arial", 12, "bold")).pack(pady=(10, 5))
        ttkb.Label(self, text="Use as variáveis entre chaves { } para que o robô substitua pelos dados reais.", font=("Arial", 9)).pack(pady=0)
        
        info_vars = (
            "Variáveis disponíveis:\n"
            "{nome} - Nome do Aluno\n"
            "{parceiro} - Instituição (SENAI/SENAC)\n"
            "{curso} - Nome do Curso\n"
            "{idade_minima} - Idade Mínima\n"
            "{duracao} - Data de Início/Fim\n"
            "{horario} - Horário do Curso"
        )
        lbl_info = ttkb.Label(self, text=info_vars, bootstyle="info", justify="left")
        lbl_info.pack(pady=10)

        btn_salvar = ttkb.Button(self, text="💾 Salvar Mensagem", command=self.salvar, bootstyle="success")
        btn_salvar.pack(side="bottom", pady=10)

        self.txt_mensagem = scrolledtext.ScrolledText(self, width=80, height=15, font=("Arial", 10))
        self.txt_mensagem.pack(padx=10, pady=5, expand=True, fill="both")
        
        texto_atual = carregar_mensagem_padrao()
        self.txt_mensagem.insert("1.0", texto_atual)

    def salvar(self):
        novo_texto = self.txt_mensagem.get("1.0", tk.END).strip()
        salvar_mensagem_padrao(novo_texto)
        tkinter.messagebox.showinfo("Sucesso", "Mensagem padrão atualizada com sucesso!")
        self.destroy()

class CourseEditor(ttkb.Toplevel):
    def __init__(self, parent, dados_atuais, callback_salvar):
        super().__init__(parent)
        self.title("Gerenciador de Cursos e Categorias")
        self.geometry("600x450")
        self.dados = dados_atuais 
        self.callback_salvar = callback_salvar 

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        ttkb.Label(self, text="1. Categorias (Grupos)", font=("Arial", 10, "bold")).grid(row=0, column=0, pady=10)
        
        self.listbox_categorias = tk.Listbox(self, exportselection=False)
        self.listbox_categorias.grid(row=1, column=0, sticky="nsew", padx=10)
        self.listbox_categorias.bind("<<ListboxSelect>>", self.ao_selecionar_categoria)

        frame_botoes_cat = ttkb.Frame(self)
        frame_botoes_cat.grid(row=2, column=0, pady=5)
        ttkb.Button(frame_botoes_cat, text="+ Add Categoria", command=self.add_categoria, bootstyle="success-outline", width=15).pack(pady=2)
        ttkb.Button(frame_botoes_cat, text="- Remover", command=self.del_categoria, bootstyle="danger-outline", width=15).pack(pady=2)

        ttkb.Label(self, text="2. Cursos da Categoria", font=("Arial", 10, "bold")).grid(row=0, column=1, pady=10)
        
        self.listbox_cursos = tk.Listbox(self, exportselection=False)
        self.listbox_cursos.grid(row=1, column=1, sticky="nsew", padx=10)

        frame_botoes_cur = ttkb.Frame(self)
        frame_botoes_cur.grid(row=2, column=1, pady=5)
        ttkb.Button(frame_botoes_cur, text="+ Add Curso", command=self.add_curso, bootstyle="info-outline", width=15).pack(pady=2)
        ttkb.Button(frame_botoes_cur, text="- Remover", command=self.del_curso, bootstyle="danger-outline", width=15).pack(pady=2)

        ttkb.Button(self, text="💾 SALVAR ALTERAÇÕES", command=self.salvar_e_fechar, bootstyle="success").grid(row=3, column=0, columnspan=2, pady=15, sticky="ew", padx=20)

        self.atualizar_lista_categorias()

    def atualizar_lista_categorias(self):
        self.listbox_categorias.delete(0, tk.END)
        for cat in self.dados.keys():
            self.listbox_categorias.insert(tk.END, cat)

    def ao_selecionar_categoria(self, event):
        selection = self.listbox_categorias.curselection()
        if selection:
            cat = self.listbox_categorias.get(selection[0])
            self.atualizar_lista_cursos(cat)

    def atualizar_lista_cursos(self, categoria):
        self.listbox_cursos.delete(0, tk.END)
        cursos = self.dados.get(categoria, [])
        for curso in cursos:
            self.listbox_cursos.insert(tk.END, curso)

    def add_categoria(self):
        nova_cat = simpledialog.askstring("Nova Categoria", "Nome da nova categoria (ex: Marketing):", parent=self)
        if nova_cat and nova_cat not in self.dados:
            self.dados[nova_cat] = []
            self.atualizar_lista_categorias()

    def del_categoria(self):
        selection = self.listbox_categorias.curselection()
        if selection:
            cat = self.listbox_categorias.get(selection[0])
            if tkinter.messagebox.askyesno("Confirmar", f"Tem certeza que deseja apagar a categoria '{cat}' e todos os seus cursos?"):
                del self.dados[cat]
                self.atualizar_lista_categorias()
                self.listbox_cursos.delete(0, tk.END)

    def add_curso(self):
        selection = self.listbox_categorias.curselection()
        if not selection:
            tkinter.messagebox.showwarning("Aviso", "Selecione uma categoria primeiro!")
            return
        
        cat = self.listbox_categorias.get(selection[0])
        novo_curso = simpledialog.askstring("Novo Curso", f"Nome do curso para adicionar em '{cat}':", parent=self)
        
        if novo_curso:
            self.dados[cat].append(novo_curso)
            self.atualizar_lista_cursos(cat)

    def del_curso(self):
        sel_cat = self.listbox_categorias.curselection()
        sel_cur = self.listbox_cursos.curselection()
        
        if sel_cat and sel_cur:
            cat = self.listbox_categorias.get(sel_cat[0])
            curso = self.listbox_cursos.get(sel_cur[0])
            
            self.dados[cat].remove(curso)
            self.atualizar_lista_cursos(cat)

    def salvar_e_fechar(self):
        self.callback_salvar(self.dados)
        self.destroy()

class CourseOfferGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Disparador de Mensagens")
        self.root.geometry("550x900") 

        self.settings = load_settings()
        self.style = Style(theme=self.settings["theme"])
        self.last_line = load_last_line()
        self.numeros_enviados = carregar_numeros_enviados()
        self.config_cursos = carregar_cursos()

        header_frame = ttk.Frame(root)
        header_frame.pack(fill='x', padx=10, pady=10)

        self.btn_edit_courses = ttkb.Button(header_frame, text="⚙️ Gerenciar Cursos", command=self.abrir_editor_cursos, bootstyle="secondary-outline")
        self.btn_edit_courses.pack(side="left", padx=5)

        self.btn_edit_message = ttkb.Button(header_frame, text="📝 Editar Mensagem", command=self.abrir_editor_mensagem, bootstyle="secondary-outline")
        self.btn_edit_message.pack(side="left", padx=5)

        self.menu_theme = ttk.Menubutton(header_frame, text="Escolher Tema")
        self.theme = tk.Menu(self.menu_theme, tearoff=-1)

        themes = ["superhero", "flatly", "darkly", "journal", "cyborg", "lumen", "minty", "pulse", "sandstone", "solar", "united", "yeti", "cerulean", "cosmo", "litera", "morph", "simplex", "vapor"]

        for theme in themes:
            self.theme.add_command(label=theme.capitalize(), command=lambda t=theme: self.change_theme(t))

        self.menu_theme.configure(menu=self.theme)
        self.menu_theme.pack(side="left", padx=5)

        self.title_label = ttk.Label(self.root, text='Disparador de Mensagens', font=('Arial', 20))
        self.title_label.pack(pady=5)

        self.caminho_arquivo = tk.StringVar()
        
        self.btn_selecionar_arquivo = ttkb.Button(self.root, text="📁 Selecionar Planilha de Alunos", command=self.selecionar_arquivo, bootstyle='info-outline')
        self.btn_selecionar_arquivo.pack(pady=5)

        self.label_arquivo = ttk.Label(self.root, textvariable=self.caminho_arquivo, font=('Arial', 8), foreground="gray")
        self.label_arquivo.pack(pady=2)

        # Checkbox para envio personalizado
        self.simple_mode_var = tk.BooleanVar(value=False)
        self.chk_simple_mode = ttkb.Checkbutton(
            root, 
            text="Enviar Apenas Mensagem Personalizada (Ignorar Variáveis)", 
            variable=self.simple_mode_var,
            command=self.toggle_inputs, 
            bootstyle="round-toggle"
        )
        self.chk_simple_mode.pack(pady=10)

        self.menu_course = ttk.Menubutton(root, text="Selecione um Curso")
        self.course_selected = tk.StringVar()
        self.course = tk.Menu(self.menu_course, tearoff=0)

        self.atualizar_menu_cursos()

        self.menu_course.configure(menu=self.course)
        self.menu_course.pack(pady=5)

        self.menu_button = ttk.Menubutton(root, text="Instituição Parceira")
        self.partner_selected = tk.StringVar()
        self.parceiro = tk.Menu(self.menu_button, tearoff=0)

        self.parceiro.add_command(label="SENAC", command=lambda: self.Chosing_Partner("SENAC"))
        self.parceiro.add_command(label="SENAI", command=lambda: self.Chosing_Partner("SENAI"))

        self.menu_button.configure(menu=self.parceiro)
        self.menu_button.pack(pady=5)

        frame_group = ttk.Frame(root)
        frame_group.pack(pady=5, anchor="center") 

        bg_color = self.style.lookup("TFrame", "background") 
        lbl_spacer_group = ttk.Label(frame_group, text="(?)", font=("Arial", 9, "bold"), foreground=bg_color)
        lbl_spacer_group.pack(side="left", padx=5)

        self.menu_group = ttk.Menubutton(frame_group, text="Deseja enviar mensagem por grupos?")
        self.group_selected = tk.StringVar()
        self.group = tk.Menu(self.menu_group, tearoff=0)
        self.group.add_command(label="SIM", command=lambda: self.Chosing_Group("SIM"))
        self.group.add_command(label="NÃO", command=lambda: self.Chosing_Group("NÃO"))
        self.menu_group.configure(menu=self.group)
        self.menu_group.pack(side="left")

        lbl_help_group = ttk.Label(frame_group, text="(?)", font=("Arial", 9, "bold"), foreground="#17a2b8", cursor="hand2")
        lbl_help_group.pack(side="left", padx=5)
        
        texto_ajuda_grupo = (
            "COMO FUNCIONA O ENVIO POR GRUPOS:\n\n"
            "• SIM: O robô enviará mensagem para todos os alunos que escolheram cursos\n"
            "da mesma CATEGORIA do curso selecionado (Ex: Cursos de TI, Cursos de Marketing).\n"
            "Use isso se o curso for genérico ou interessar a uma área inteira.\n\n"
            "• NÃO: O robô enviará mensagem APENAS para quem escolheu EXATAMENTE\n"
            "o nome do curso que você selecionou no menu."
        )
        create_tooltip(lbl_help_group, texto_ajuda_grupo)

        self.schedule_label = ttk.Label(root, text="Horário:")
        self.schedule_label.pack()
        self.schedule_entry = ttk.Entry(self.root, width=30, bootstyle='info')
        self.schedule_entry.pack()
        add_placeholder(self.schedule_entry, "Ex: 19:00 às 22:00")

        self.minage_label = ttk.Label(root, text="Idade mínima:")
        self.minage_label.pack()
        self.minage_entry = ttk.Entry(root, width=30, bootstyle='info')
        self.minage_entry.pack()

        self.duration_label = ttk.Label(root, text="Data de Início e Fim:")
        self.duration_label.pack()
        self.duration_entry = ttk.Entry(root, width=30, bootstyle='info')
        self.duration_entry.pack()
        add_placeholder(self.duration_entry, "Ex: 10/02 a 15/02")

        self.minrange_label = ttk.Label(root, text="De qual linha devo começar:")
        self.minrange_label.pack()
        self.minrange_entry = ttk.Entry(root, width=30, bootstyle='info')
        self.minrange_entry.pack()
        add_placeholder(self.minrange_entry, f"Última linha enviada: {self.last_line}")

        self.maxrange_label = ttk.Label(root, text="Até qual linha devo enviar:")
        self.maxrange_label.pack()
        self.maxrange_entry = ttk.Entry(root, width=30, bootstyle='info')
        self.maxrange_entry.pack()

        self.send_button = ttkb.Button(self.root, text="Enviar Mensagens", command=self.start_sending, bootstyle='success')
        self.send_button.pack(pady=10)

        self.progress = ttkb.Progressbar(self.root, orient="horizontal", length=300, mode="determinate", bootstyle="success-striped")
        self.progress.pack(pady=5)

        self.cancel_button = ttkb.Button(self.root, text="Interromper Envio", command=self.interromper_codigo, bootstyle='danger')
        self.cancel_button.pack(pady=5)

        frame_clean = ttk.Frame(root)
        frame_clean.pack(pady=10, anchor="center")

        bg_color = self.style.lookup("TFrame", "background")
        lbl_spacer_clean = ttk.Label(frame_clean, text="(?)", font=("Arial", 9, "bold"), foreground=bg_color)
        lbl_spacer_clean.pack(side="left", padx=5)

        self.clean_button = ttkb.Button(frame_clean, text="Limpar Histórico de Envios", command=self.limpar_historico_numeros, bootstyle='warning-outline')
        self.clean_button.pack(side="left")

        lbl_help_clean = ttk.Label(frame_clean, text="(?)", font=("Arial", 9, "bold"), foreground="#ffc107", cursor="hand2")
        lbl_help_clean.pack(side="left", padx=5)

        texto_ajuda_limpar = (
            "CUIDADO: ESTA AÇÃO É IRREVERSÍVEL!\n\n"
            "O robô salva uma lista de telefones que já receberam mensagens para não enviar duplicado.\n"
            "Use este botão APENAS se você for iniciar uma NOVA campanha de divulgação\n"
            "e quiser que as pessoas que já receberam mensagens no passado possam receber novamente."
        )
        create_tooltip(lbl_help_clean, texto_ajuda_limpar)

        self.credits_label = ttkb.Label(self.root, text='developed by: Lucas Ferrari, Eduardo Zanin & João Miguel.', font=('Arial', 7, 'bold'))
        self.credits_label.pack(pady=1)

        self.running = False

    def toggle_inputs(self):
        state = "disabled" if self.simple_mode_var.get() else "normal"
        
        self.schedule_entry.config(state=state)
        self.minage_entry.config(state=state)
        self.duration_entry.config(state=state)
        self.menu_button.config(state=state)
        self.menu_group.config(state=state)
        self.menu_course.config(state=state)

    def abrir_editor_cursos(self):
        CourseEditor(self.root, self.config_cursos, self.salvar_alteracoes_cursos)

    def abrir_editor_mensagem(self):
        MessageEditor(self.root)

    def salvar_alteracoes_cursos(self, novos_dados):
        self.config_cursos = novos_dados
        salvar_cursos_json(self.config_cursos)
        self.atualizar_menu_cursos() 
        tkinter.messagebox.showinfo("Sucesso", "Lista de cursos atualizada com sucesso!")

    def atualizar_menu_cursos(self):
        self.course.delete(0, tk.END) 
        for categoria, cursos in self.config_cursos.items():
            for curso in cursos:
                self.course.add_command(label=curso, command=lambda c=curso: self.Chosing_Course(c))
    
    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione a planilha de alunos",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if arquivo:
            self.caminho_arquivo.set(arquivo)

    def limpar_historico_numeros(self):
        if self.running:
            tkinter.messagebox.showwarning("Aviso", "Pare o envio de mensagens antes de limpar o histórico.")
            return

        resposta = tkinter.messagebox.askyesno("Confirmar Limpeza", 
                                               "ATENÇÃO: Isso apagará a lista de todos os números que já receberam mensagens.\n\n"
                                               "Se você fizer isso, o robô poderá enviar mensagens repetidas para pessoas que já receberam anteriormente.\n\n"
                                               "Deseja continuar?")
        if resposta:
            self.numeros_enviados = set() 
            salvar_numeros_enviados(self.numeros_enviados) 
            tkinter.messagebox.showinfo("Sucesso", "Histórico de envios limpo com sucesso! Agora você pode enviar novas mensagens para os mesmos números.")

    def interromper_codigo(self):
        print("Envio encerrado")
        self.running = False

    def start_sending(self):
        if self.running:
            print("O envio de mensagens já está em andamento.")
            return
            
        if not self.caminho_arquivo.get():
            tkinter.messagebox.showwarning("Aviso", "Por favor, selecione a planilha de alunos primeiro.")
            return

        if not self.simple_mode_var.get():
            horario = self.schedule_entry.get()
            duracao = self.duration_entry.get()

            if "Ex:" in horario or not horario.strip():
                tkinter.messagebox.showwarning("Atenção", "Por favor, preencha o campo HORÁRIO corretamente.")
                return
            
            if "Ex:" in duracao or not duracao.strip():
                tkinter.messagebox.showwarning("Atenção", "Por favor, preencha o campo DATA DE INÍCIO E FIM corretamente.")
                return

        self.running = True
        print('Iniciado o envio de mensagens!')
        t = threading.Thread(target=self.send_messages)
        t.start()

    def Chosing_Course(self, curso: str):
        self.menu_course.config(text=curso)
        self.course_selected.set(curso)

    def Chosing_Partner(self, parceiro: str):
        self.menu_button.config(text=parceiro)
        self.partner_selected.set(parceiro)

    def Chosing_Group(self, grupo: str):
        self.menu_group.config(text=grupo)
        self.group_selected.set(grupo)

    def change_theme(self, theme):
        self.style.theme_use(theme)
        self.settings["theme"] = theme
        save_settings(self.settings)

    def encontra_categoria(self, curso_de_envio: str) -> list:
        for categoria, lista_cursos in self.config_cursos.items():
            if curso_de_envio in lista_cursos:
                return lista_cursos
        return []

    def send_messages(self):
        try:
            curso_de_envio = self.course_selected.get()
            
            # Se for modo simples, variáveis ficam vazias e evitamos erro de conversão
            if self.simple_mode_var.get():
                parceiro = ""
                horario_do_curso = ""
                data_de_duracao = ""
                idademin = 0 # Valor seguro para evitar erro de int() com string vazia
                por_grupo = self.group_selected.get()
            else:
                parceiro = self.partner_selected.get()
                horario_do_curso = self.schedule_entry.get()
                data_de_duracao = self.duration_entry.get()
                
                # Validação segura para idade mínima
                entrada_idade = self.minage_entry.get()
                if entrada_idade and entrada_idade.isdigit():
                    idademin = int(entrada_idade)
                else:
                    idademin = 0
                    
                por_grupo = self.group_selected.get()   

            entrada_min = self.minrange_entry.get()
            if "Última linha enviada:" in entrada_min:
                try:
                    linhamin = int(entrada_min.split(": ")[1])
                except (IndexError, ValueError):
                    linhamin = self.last_line
            else:
                linhamin = int(entrada_min)

            linhamax = int(self.maxrange_entry.get())
            
            caminho_planilha = self.caminho_arquivo.get()
            
            try:
                alunos = pd.read_excel(caminho_planilha)
            except PermissionError:
                tkinter.messagebox.showerror("Erro", "O arquivo Excel parece estar aberto. Por favor, feche-o e tente novamente.")
                self.running = False
                return
            except Exception as e:
                tkinter.messagebox.showerror("Erro", f"Erro ao ler a planilha: {e}")
                self.running = False
                return

            numeros_enviados = self.numeros_enviados
            self.running = True

            total_alunos = linhamax - linhamin
            ultima_linha_enviada = None
            
            self.progress['maximum'] = total_alunos
            self.progress['value'] = 0

            for x in range(linhamin, linhamax):
                if not self.running:
                    print('Código interrompido na linha: {0}'.format(x))
                    break
                
                self.progress['value'] += 1
                self.root.update_idletasks()

                try:
                    linha_correta = x + 2
                    
                    modelo_mensagem = carregar_mensagem_padrao()

                    if self.simple_mode_var.get():
                        nome = alunos.loc[x, 'Nome Completo']
                        telefone = limpar_telefone(alunos.loc[x, "Whatsapp com DDD (somente números - sem espaço)"])

                        if telefone is None:
                            print(f"Telefone inválido na linha {linha_correta}. Pulando...")
                            continue

                        if telefone in self.numeros_enviados:
                            continue

                        mensagem = modelo_mensagem 
                        
                        link_mensagem_whatsapp = f'https://web.whatsapp.com/send/?phone={telefone}&text={quote(mensagem)}'
                        webbrowser.open(link_mensagem_whatsapp)
                        sleep(2)
                        sleep(6)
                        pyautogui.press('enter')
                        sleep(6)
                        pyautogui.hotkey('ctrl', 'w')

                        self.numeros_enviados.add(telefone)
                        salvar_numeros_enviados(self.numeros_enviados)

                        logging.info(f'Messagem GENÉRICA enviada para: {nome}, Telefone: {telefone}, Linha: {linha_correta}')
                        save_last_line(linha_correta)
                    
                    else:
                        cursos = alunos.loc[x, "Dentre as opções qual curso gostaria de fazer?"]
                        if pd.isna(cursos):
                            continue
                            
                        lista_cursos = cursos.split(sep=', ')

                        if por_grupo == "SIM":
                            categoria = self.encontra_categoria(curso_de_envio)
                            for curso in lista_cursos:
                                if curso in categoria:
                                    nome = alunos.loc[x, 'Nome Completo']
                                    telefone = limpar_telefone(alunos.loc[x, "Whatsapp com DDD (somente números - sem espaço)"])

                                    if telefone is None:
                                        print(f"Telefone inválido na linha {linha_correta}. Pulando...")
                                        continue

                                    if telefone in self.numeros_enviados:
                                        continue

                                    mensagem = modelo_mensagem.format(
                                        nome=nome,
                                        parceiro=parceiro,
                                        curso=curso_de_envio,
                                        idade_minima=idademin,
                                        duracao=data_de_duracao,
                                        horario=horario_do_curso
                                    )
                                    
                                    link_mensagem_whatsapp = f'https://web.whatsapp.com/send/?phone={telefone}&text={quote(mensagem)}'
                                    webbrowser.open(link_mensagem_whatsapp)
                                    sleep(2)
                                    sleep(6)
                                    pyautogui.press('enter')
                                    sleep(6)
                                    pyautogui.hotkey('ctrl', 'w')

                                    self.numeros_enviados.add(telefone)
                                    salvar_numeros_enviados(self.numeros_enviados)

                                    logging.info(f'Messagem enviada para: {nome}, Telefone: {telefone}, Curso: {curso_de_envio}, Linha: {linha_correta}')
                                    save_last_line(linha_correta)

                        else:
                            for curso in lista_cursos:
                                if curso.upper() == curso_de_envio.upper():
                                    nome = alunos.loc[x, 'Nome Completo']
                                    telefone = limpar_telefone(alunos.loc[x, 'Whatsapp com DDD (somente números - sem espaço)'])

                                    if telefone is None:
                                        print(f"Telefone inválido na linha {linha_correta}. Pulando...")
                                        continue

                                    if telefone in numeros_enviados:
                                        continue

                                    mensagem = modelo_mensagem.format(
                                        nome=nome,
                                        parceiro=parceiro,
                                        curso=curso_de_envio,
                                        idade_minima=idademin,
                                        duracao=data_de_duracao,
                                        horario=horario_do_curso
                                    )

                                    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
                                    webbrowser.open(link_mensagem_whatsapp)
                                    sleep(2)
                                    sleep(6)
                                    pyautogui.press('enter')
                                    sleep(6)
                                    pyautogui.hotkey('ctrl', 'w')

                                    self.numeros_enviados.add(telefone)
                                    salvar_numeros_enviados(self.numeros_enviados)

                                    logging.info(f'Messagem enviada para: {nome}, Telefone: {telefone}, Curso: {curso_de_envio}, Linha: {linha_correta}')
                                    save_last_line(linha_correta)

                except Exception as e:
                    print(f"Erro ao processar linha {x}: {e}")
                    continue

            print("Todas as linhas foram lidas!")
            self.running = False
            self.progress['value'] = 0

            ttk.Label(self.root, text="Envio de mensagens concluído!", foreground="green").pack(pady=10)

        except Exception as e:
            print(f"Erro no envio de mensagens: {e}")
            tkinter.messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

root = tk.Tk()
gui = CourseOfferGUI(root)
root.mainloop()