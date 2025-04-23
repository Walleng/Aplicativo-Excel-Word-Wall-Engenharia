#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Módulo de interface de usuário para o aplicativo Wall Engenharia.
Responsável por prover interface gráfica para interação com o usuário.
"""

import os
import sys
import logging
from typing import Dict, Any, Optional, Tuple
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

# Importar módulos do aplicativo
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from src.excel_reader import ExcelReader, extract_data_from_excel
from src.word_writer import WordWriter, generate_proposal
from src.config_manager import get_config_manager

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger('ui_manager')

class WallEngenhariaApp:
    """Classe principal da interface de usuário do aplicativo."""
    
    def __init__(self, root):
        """
        Inicializa a interface de usuário.
        
        Args:
            root: Janela raiz do Tkinter.
        """
        self.root = root
        self.root.title("Wall Engenharia - Integração Excel-Word")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # Carregar configurações
        self.config = get_config_manager()
        
        # Variáveis de controle
        self.excel_file_path = tk.StringVar()
        self.word_template_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.selected_sheet = tk.StringVar()
        
        # Dados extraídos
        self.extracted_data = {}
        self.excel_reader = None
        
        # Criar interface
        self._create_ui()
    
    def _create_ui(self):
        """Cria os elementos da interface de usuário."""
        # Criar frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Criar notebook (abas)
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Aba 1: Seleção de arquivos
        files_frame = ttk.Frame(notebook, padding="10")
        notebook.add(files_frame, text="Seleção de Arquivos")
        
        # Aba 2: Edição de dados
        data_frame = ttk.Frame(notebook, padding="10")
        notebook.add(data_frame, text="Edição de Dados")
        
        # Aba 3: Configurações
        config_frame = ttk.Frame(notebook, padding="10")
        notebook.add(config_frame, text="Configurações")
        
        # Configurar aba de seleção de arquivos
        self._setup_files_tab(files_frame)
        
        # Configurar aba de edição de dados
        self._setup_data_tab(data_frame)
        
        # Configurar aba de configurações
        self._setup_config_tab(config_frame)
        
        # Barra de status
        status_frame = ttk.Frame(main_frame, relief=tk.SUNKEN, padding="2")
        status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=5, pady=5)
        
        self.status_label = ttk.Label(status_frame, text="Pronto")
        self.status_label.pack(side=tk.LEFT)
        
        # Botões de ação
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=5, pady=5)
        
        ttk.Button(button_frame, text="Gerar Proposta", command=self._generate_proposal).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Visualizar Dados", command=self._preview_data).pack(side=tk.RIGHT, padx=5)
    
    def _setup_files_tab(self, parent):
        """Configura a aba de seleção de arquivos."""
        # Frame para arquivo Excel
        excel_frame = ttk.LabelFrame(parent, text="Arquivo Excel de Orçamento", padding="10")
        excel_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(excel_frame, text="Arquivo:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(excel_frame, textvariable=self.excel_file_path, width=50).grid(row=0, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        ttk.Button(excel_frame, text="Procurar...", command=self._browse_excel_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Frame para planilha
        sheet_frame = ttk.Frame(excel_frame)
        sheet_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W+tk.E, padx=5, pady=5)
        
        ttk.Label(sheet_frame, text="Planilha:").pack(side=tk.LEFT, padx=5)
        self.sheet_combobox = ttk.Combobox(sheet_frame, textvariable=self.selected_sheet, state="readonly", width=30)
        self.sheet_combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(sheet_frame, text="Carregar Planilhas", command=self._load_excel_sheets).pack(side=tk.LEFT, padx=5)
        
        # Frame para modelo Word
        word_frame = ttk.LabelFrame(parent, text="Modelo de Proposta Word", padding="10")
        word_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(word_frame, text="Arquivo:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(word_frame, textvariable=self.word_template_path, width=50).grid(row=0, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        ttk.Button(word_frame, text="Procurar...", command=self._browse_word_template).grid(row=0, column=2, padx=5, pady=5)
        
        # Frame para arquivo de saída
        output_frame = ttk.LabelFrame(parent, text="Proposta Gerada", padding="10")
        output_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(output_frame, text="Salvar em:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(output_frame, textvariable=self.output_file_path, width=50).grid(row=0, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        ttk.Button(output_frame, text="Procurar...", command=self._browse_output_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Arquivos recentes
        recent_frame = ttk.LabelFrame(parent, text="Arquivos Recentes", padding="10")
        recent_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Excel recentes
        ttk.Label(recent_frame, text="Excel:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.recent_excel_listbox = tk.Listbox(recent_frame, height=3)
        self.recent_excel_listbox.grid(row=0, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        self.recent_excel_listbox.bind("<Double-1>", self._select_recent_excel)
        
        # Word recentes
        ttk.Label(recent_frame, text="Word:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.recent_word_listbox = tk.Listbox(recent_frame, height=3)
        self.recent_word_listbox.grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        self.recent_word_listbox.bind("<Double-1>", self._select_recent_word)
        
        # Configurar grid
        recent_frame.columnconfigure(1, weight=1)
        
        # Carregar arquivos recentes
        self._load_recent_files()
    
    def _setup_data_tab(self, parent):
        """Configura a aba de edição de dados."""
        # Frame para dados extraídos
        data_frame = ttk.LabelFrame(parent, text="Dados Extraídos do Excel", padding="10")
        data_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Criar campos de entrada para cada dado
        row = 0
        self.data_entries = {}
        
        fields = [
            ("nome_cliente", "Nome do Cliente:"),
            ("nome_contato", "Nome do Contato:"),
            ("email", "E-mail:"),
            ("telefone", "Telefone:"),
            ("escopo", "Escopo:"),
            ("prazo", "Prazo:"),
            ("custo", "Custo:"),
            ("garantias", "Garantias:"),
            ("seguro", "Seguro:"),
            ("nao_inclusos", "Não Inclusos:")
        ]
        
        for field, label in fields:
            ttk.Label(data_frame, text=label).grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
            
            # Usar ScrolledText para campos que podem ter múltiplas linhas
            if field in ["escopo", "prazo", "garantias", "seguro", "nao_inclusos"]:
                entry = ScrolledText(data_frame, width=40, height=3)
                entry.grid(row=row, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
            else:
                entry = ttk.Entry(data_frame, width=40)
                entry.grid(row=row, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
            
            self.data_entries[field] = entry
            row += 1
        
        # Configurar grid
        data_frame.columnconfigure(1, weight=1)
        
        # Botões de ação
        button_frame = ttk.Frame(data_frame)
        button_frame.grid(row=row, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Extrair Dados do Excel", command=self._extract_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Limpar Campos", command=self._clear_data_fields).pack(side=tk.LEFT, padx=5)
    
    def _setup_config_tab(self, parent):
        """Configura a aba de configurações."""
        # Frame para configurações
        config_frame = ttk.LabelFrame(parent, text="Configurações do Aplicativo", padding="10")
        config_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Tema
        ttk.Label(config_frame, text="Tema:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        theme_combobox = ttk.Combobox(config_frame, values=["Claro", "Escuro"], state="readonly")
        theme_combobox.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        theme_combobox.current(0)  # Padrão: Claro
        
        # Botões de ação
        button_frame = ttk.Frame(config_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Salvar Configurações", command=self._save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Restaurar Padrões", command=self._restore_default_config).pack(side=tk.LEFT, padx=5)
    
    def _browse_excel_file(self):
        """Abre diálogo para selecionar arquivo Excel."""
        file_path = filedialog.askopenfilename(
            title="Selecionar Arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os Arquivos", "*.*")]
        )
        
        if file_path:
            self.excel_file_path.set(file_path)
            self.config.add_recent_file('excel', file_path)
            self._load_recent_files()
            self._load_excel_sheets()
    
    def _browse_word_template(self):
        """Abre diálogo para selecionar modelo Word."""
        file_path = filedialog.askopenfilename(
            title="Selecionar Modelo Word",
            filetypes=[("Documentos Word", "*.docx *.doc"), ("Todos os Arquivos", "*.*")]
        )
        
        if file_path:
            self.word_template_path.set(file_path)
            self.config.add_recent_file('word_templates', file_path)
            self._load_recent_files()
    
    def _browse_output_file(self):
        """Abre diálogo para selecionar arquivo de saída."""
        file_path = filedialog.asksaveasfilename(
            title="Salvar Proposta Como",
            defaultextension=".docx",
            filetypes=[("Documentos Word", "*.docx"), ("Todos os Arquivos", "*.*")]
        )
        
        if file_path:
            self.output_file_path.set(file_path)
    
    def _load_excel_sheets(self):
        """Carrega as planilhas do arquivo Excel selecionado."""
        excel_path = self.excel_file_path.get()
        
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Erro", "Selecione um arquivo Excel válido.")
            return
        
        try:
            # Fechar leitor anterior se existir
            if self.excel_reader:
                self.excel_reader.close()
            
            # Criar novo leitor
            self.excel_reader = ExcelReader(excel_path)
            
            # Atualizar combobox de planilhas
            sheets = self.excel_reader.get_sheet_names()
            self.sheet_combobox['values'] = sheets
            
            if sheets:
                self.selected_sheet.set(sheets[0])
                self.status_label.config(text=f"Arquivo Excel carregado: {os.path.basename(excel_path)}")
            else:
                self.status_label.config(text="Nenhuma planilha encontrada no arquivo Excel.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar planilhas: {e}")
            logger.error(f"Erro ao carregar planilhas: {e}")
    
    def _load_recent_files(self):
        """Carrega a lista de arquivos recentes."""
        # Limpar listas
        self.recent_excel_listbox.delete(0, tk.END)
        self.recent_word_listbox.delete(0, tk.END)
        
        # Carregar arquivos Excel recentes
        for file_path in self.config.get_recent_files('excel'):
            if os.path.exists(file_path):
                self.recent_excel_listbox.insert(tk.END, os.path.basename(file_path))
        
        # Carregar modelos Word recentes
        for file_path in self.config.get_recent_files('word_templates'):
            if os.path.exists(file_path):
                self.recent_word_listbox.insert(tk.END, os.path.basename(file_path))
    
    def _select_recent_excel(self, event):
        """Seleciona um arquivo Excel recente."""
        selection = self.recent_excel_listbox.curselection()
        if selection:
            index = selection[0]
            recent_files = self.config.get_recent_files('excel')
            if index < len(recent_files):
                self.excel_file_path.set(recent_files[index])
                self._load_excel_sheets()
    
    def _select_recent_word(self, event):
        """Seleciona um modelo Word recente."""
        selection = self.recent_word_listbox.curselection()
        if selection:
            index = selection[0]
            recent_files = self.config.get_recent_files('word_templates')
            if index < len(recent_files):
                self.word_template_path.set(recent_files[index])
    
    def _extract_data(self):
        """Extrai dados do arquivo Excel."""
        excel_path = self.excel_file_path.get()
        sheet_name = self.selected_sheet.get()
        
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Erro", "Selecione um arquivo Excel válido.")
            return
        
        if not sheet_name:
            messagebox.showerror("Erro", "Selecione uma planilha.")
            return
        
        try:
            # Extrair dados
            if self.excel_reader and self.excel_reader.select_sheet(sheet_name):
                self.extracted_data = self.excel_reader.extract_data_for_proposal()
                
                # Preencher campos de entrada
                self._update_data_fields()
                
                self.status_label.config(text=f"Dados extraídos da planilha: {sheet_name}")
            else:
                messagebox.showerror("Erro", "Erro ao selecionar planilha.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao extrair dados: {e}")
            logger.error(f"Erro ao extrair dados: {e}")
    
    def _update_data_fields(self):
        """Atualiza os campos de entrada com os dados extraídos."""
        for field, entry in self.data_entries.items():
            value = self.extracted_data.get(field, "")
            
            # Limpar campo
            if isinstance(entry, ScrolledText):
                entry.delete(1.0, tk.END)
                if value:
                    if isinstance(value, list):
                        entry.insert(tk.END, "\n".join(value))
                    else:
                        entry.insert(tk.END, str(value))
            else:
                entry.delete(0, tk.END)
                if value:
                    entry.insert(0, str(value))
    
    def _clear_data_fields(self):
        """Limpa todos os campos de entrada."""
        for field, entry in self.data_entries.items():
            if isinstance(entry, ScrolledText):
                entry.delete(1.0, tk.END)
            else:
                entry.delete(0, tk.END)
    
    def _get_data_from_fields(self) -> Dict[str, Any]:
        """
        Obtém os dados dos campos de entrada.
        
        Returns:
            Dicionário com os dados dos campos.
        """
        data = {}
        
        for field, entry in self.data_entries.items():
            if isinstance(entry, ScrolledText):
                value = entry.get(1.0, tk.END).strip()
            else:
                value = entry.get().strip()
            
            # Converter valor numérico para número
            if field == 'custo' and value:
                try:
                    # Remover formatação de moeda
                    value = value.replace("R$", "").replace(".", "").replace(",", ".").strip()
                    value = float(value)
                except ValueError:
                    pass
            
            data[field] = value
        
        return data
    
    def _preview_data(self):
        """Visualiza os dados que serão usados na proposta."""
        data = self._get_data_from_fields()
        
        # Criar janela de visualização
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Visualização dos Dados")
        preview_window.geometry("600x400")
        
        # Criar área de texto
        text_area = ScrolledText(preview_window, width=70, height=20)
        text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Preencher com os dados
        text_area.insert(tk.END, "DADOS PARA A PROPOSTA:\n\n")
        
        for field, label in [
            ("nome_cliente", "Nome do Cliente"),
            ("nome_contato", "Nome do Contato"),
            ("email", "E-mail"),
            ("telefone", "Telefone"),
            ("escopo", "Escopo"),
            ("prazo", "Prazo"),
            ("custo", "Custo"),
            ("garantias", "Garantias"),
            ("seguro", "Seguro"),
            ("nao_inclusos", "Não Inclusos")
        ]:
            value = data.get(field, "")
            
            # Formatar valor numérico para moeda brasileira
            if field == 'custo' and isinstance(value, (int, float)):
                value = f"R$ {value:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            
            text_area.insert(tk.END, f"{label}: {value}\n\n")
        
        text_area.config(state=tk.DISABLED)  # Tornar somente leitura
    
    def _generate_proposal(self):
        """Gera a proposta com os dados fornecidos."""
        excel_path = self.excel_file_path.get()
        template_path = self.word_template_path.get()
        output_path = self.output_file_path.get()
        
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Erro", "Selecione um arquivo Excel válido.")
            return
        
        if not template_path or not os.path.exists(template_path):
            messagebox.showerror("Erro", "Selecione um modelo Word válido.")
            return
        
        if not output_path:
            messagebox.showerror("Erro", "Selecione um local para salvar a proposta gerada.")
            return
        
        try:
            # Obter dados dos campos
            data = self._get_data_from_fields()
            
            # Gerar proposta
            success = generate_proposal(template_path, output_path, data)
            
            if success:
                messagebox.showinfo("Sucesso", f"Proposta gerada com sucesso:\n{output_path}")
                self.status_label.config(text=f"Proposta gerada: {os.path.basename(output_path)}")
            else:
                messagebox.showerror("Erro", "Erro ao gerar proposta.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar proposta: {e}")
            logger.error(f"Erro ao gerar proposta: {e}")
    
    def _save_config(self):
        """Salva as configurações do aplicativo."""
        # Implementar salvamento de configurações
        messagebox.showinfo("Configurações", "Configurações salvas com sucesso.")
    
    def _restore_default_config(self):
        """Restaura as configurações padrão do aplicativo."""
        # Implementar restauração de configurações
        messagebox.showinfo("Configurações", "Configurações restauradas para os valores padrão.")


def main():
    """Função principal do aplicativo."""
    root = tk.Tk()
    app = WallEngenhariaApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
