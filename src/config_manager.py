#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Módulo de configuração para o aplicativo Wall Engenharia.
Responsável por gerenciar configurações e mapeamentos entre Excel e Word.
"""

import os
import json
import logging
from typing import Dict, Any, Optional

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger('config_manager')

class ConfigManager:
    """Classe para gerenciar configurações do aplicativo."""
    
    def __init__(self, config_file: str = None):
        """
        Inicializa o gerenciador de configuração.
        
        Args:
            config_file: Caminho para o arquivo de configuração (opcional).
        """
        self.config_file = config_file or os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'resources', 'config', 'config.json'
        )
        self.config = self._load_default_config()
        
        # Carregar configuração do arquivo se existir
        if os.path.exists(self.config_file):
            self._load_config()
        else:
            # Criar diretório de configuração se não existir
            config_dir = os.path.dirname(self.config_file)
            if not os.path.exists(config_dir):
                os.makedirs(config_dir)
            
            # Salvar configuração padrão
            self.save_config()
    
    def _load_default_config(self) -> Dict[str, Any]:
        """
        Carrega a configuração padrão.
        
        Returns:
            Dicionário com a configuração padrão.
        """
        return {
            "excel_mappings": {
                "nome_cliente": {
                    "search_terms": ["cliente", "empresa", "contratante"],
                    "default_position": {"row_range": [1, 10], "col_range": [1, 10]}
                },
                "nome_contato": {
                    "search_terms": ["contato", "responsável", "representante"],
                    "default_position": None
                },
                "email": {
                    "search_terms": ["email", "e-mail", "correio eletrônico"],
                    "default_position": None
                },
                "telefone": {
                    "search_terms": ["telefone", "tel", "fone", "celular"],
                    "default_position": None
                },
                "escopo": {
                    "search_terms": ["escopo", "serviço", "objeto"],
                    "default_position": None
                },
                "prazo": {
                    "search_terms": ["prazo", "duração", "período"],
                    "default_position": None
                },
                "custo": {
                    "search_terms": ["custo", "valor", "preço", "total"],
                    "default_position": {"row_range": [-20, -1], "col_range": [1, 10]}
                },
                "garantias": {
                    "search_terms": ["garantia"],
                    "default_position": None
                },
                "seguro": {
                    "search_terms": ["seguro"],
                    "default_position": None
                },
                "nao_inclusos": {
                    "search_terms": ["não incluso", "não incluído"],
                    "default_position": None
                }
            },
            "word_placeholders": {
                "nome_cliente": ["{{CLIENTE}}", "{{NOME_CLIENTE}}"],
                "nome_contato": ["{{CONTATO}}", "{{NOME_CONTATO}}"],
                "email": ["{{EMAIL}}", "{{E-MAIL}}"],
                "telefone": ["{{TELEFONE}}", "{{TEL}}"],
                "escopo": ["{{ESCOPO}}"],
                "prazo": ["{{PRAZO}}"],
                "custo": ["{{CUSTO}}", "{{VALOR}}", "{{PREÇO}}"],
                "garantias": ["{{GARANTIAS}}"],
                "seguro": ["{{SEGURO}}"],
                "nao_inclusos": ["{{NAO_INCLUSOS}}", "{{NÃO_INCLUSOS}}"]
            },
            "recent_files": {
                "excel": [],
                "word_templates": []
            },
            "ui_settings": {
                "theme": "light",
                "language": "pt_BR",
                "window_size": [800, 600]
            }
        }
    
    def _load_config(self) -> None:
        """Carrega a configuração do arquivo."""
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                loaded_config = json.load(f)
                
                # Atualizar configuração com valores carregados
                for key, value in loaded_config.items():
                    self.config[key] = value
                
                logger.info(f"Configuração carregada do arquivo: {self.config_file}")
        except Exception as e:
            logger.error(f"Erro ao carregar configuração: {e}")
    
    def save_config(self) -> bool:
        """
        Salva a configuração atual no arquivo.
        
        Returns:
            True se a configuração foi salva com sucesso, False caso contrário.
        """
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
                
            logger.info(f"Configuração salva no arquivo: {self.config_file}")
            return True
        except Exception as e:
            logger.error(f"Erro ao salvar configuração: {e}")
            return False
    
    def get_excel_mapping(self, field: str) -> Dict[str, Any]:
        """
        Obtém o mapeamento de um campo no Excel.
        
        Args:
            field: Nome do campo.
            
        Returns:
            Dicionário com o mapeamento do campo.
        """
        return self.config["excel_mappings"].get(field, {})
    
    def get_word_placeholders(self, field: str) -> list:
        """
        Obtém os placeholders de um campo no Word.
        
        Args:
            field: Nome do campo.
            
        Returns:
            Lista de placeholders para o campo.
        """
        return self.config["word_placeholders"].get(field, [])
    
    def add_recent_file(self, file_type: str, file_path: str, max_files: int = 10) -> None:
        """
        Adiciona um arquivo à lista de arquivos recentes.
        
        Args:
            file_type: Tipo de arquivo ('excel' ou 'word_templates').
            file_path: Caminho do arquivo.
            max_files: Número máximo de arquivos recentes a manter.
        """
        if file_type in self.config["recent_files"]:
            # Remover o arquivo da lista se já existir
            if file_path in self.config["recent_files"][file_type]:
                self.config["recent_files"][file_type].remove(file_path)
            
            # Adicionar o arquivo no início da lista
            self.config["recent_files"][file_type].insert(0, file_path)
            
            # Limitar o número de arquivos recentes
            self.config["recent_files"][file_type] = self.config["recent_files"][file_type][:max_files]
            
            # Salvar a configuração
            self.save_config()
    
    def get_recent_files(self, file_type: str) -> list:
        """
        Obtém a lista de arquivos recentes.
        
        Args:
            file_type: Tipo de arquivo ('excel' ou 'word_templates').
            
        Returns:
            Lista de caminhos de arquivos recentes.
        """
        return self.config["recent_files"].get(file_type, [])
    
    def get_ui_setting(self, setting: str) -> Any:
        """
        Obtém uma configuração da interface de usuário.
        
        Args:
            setting: Nome da configuração.
            
        Returns:
            Valor da configuração.
        """
        return self.config["ui_settings"].get(setting)
    
    def set_ui_setting(self, setting: str, value: Any) -> None:
        """
        Define uma configuração da interface de usuário.
        
        Args:
            setting: Nome da configuração.
            value: Valor da configuração.
        """
        if setting in self.config["ui_settings"]:
            self.config["ui_settings"][setting] = value
            self.save_config()


# Instância global do gerenciador de configuração
config_manager = None

def get_config_manager(config_file: Optional[str] = None) -> ConfigManager:
    """
    Obtém a instância global do gerenciador de configuração.
    
    Args:
        config_file: Caminho para o arquivo de configuração (opcional).
        
    Returns:
        Instância do gerenciador de configuração.
    """
    global config_manager
    
    if config_manager is None:
        config_manager = ConfigManager(config_file)
    
    return config_manager


if __name__ == "__main__":
    # Teste básico do módulo
    config = get_config_manager()
    print("Configuração carregada:")
    print(json.dumps(config.config, ensure_ascii=False, indent=2))
    
    # Adicionar arquivo recente de teste
    config.add_recent_file('excel', 'C:/Exemplo/planilha.xlsx')
    print("\nArquivos Excel recentes:", config.get_recent_files('excel'))
