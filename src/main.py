#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Módulo principal do aplicativo Wall Engenharia.
Ponto de entrada da aplicação.
"""

import os
import sys
import logging
import tkinter as tk

# Adicionar diretório pai ao path para importar módulos
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Importar módulos do aplicativo
from src.ui_manager import WallEngenhariaApp

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'wall_app.log')),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger('main')

def main():
    """Função principal do aplicativo."""
    try:
        # Criar diretório de configuração se não existir
        config_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'resources', 'config')
        if not os.path.exists(config_dir):
            os.makedirs(config_dir)
        
        # Iniciar aplicação
        logger.info("Iniciando aplicativo Wall Engenharia")
        root = tk.Tk()
        app = WallEngenhariaApp(root)
        root.mainloop()
    except Exception as e:
        logger.error(f"Erro ao iniciar aplicativo: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()
