#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script para empacotar o aplicativo Wall Engenharia em um executável Windows.
"""

import os
import sys
import shutil
import subprocess
import platform

def check_requirements():
    """Verifica se os requisitos para empacotamento estão instalados."""
    try:
        import PyInstaller
        print("PyInstaller já está instalado.")
    except ImportError:
        print("Instalando PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # Verificar outras dependências
    dependencies = ["openpyxl", "python-docx", "pandas"]
    for dep in dependencies:
        try:
            __import__(dep)
            print(f"{dep} já está instalado.")
        except ImportError:
            print(f"Instalando {dep}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", dep])

def create_executable():
    """Cria o executável do aplicativo."""
    # Diretório base do projeto
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    # Diretório de saída para o executável
    dist_dir = os.path.join(base_dir, "dist")
    if not os.path.exists(dist_dir):
        os.makedirs(dist_dir)
    
    # Caminho para o script principal
    main_script = os.path.join(base_dir, "src", "main.py")
    
    # Verificar se o script principal existe
    if not os.path.exists(main_script):
        print(f"Erro: Script principal não encontrado: {main_script}")
        return False
    
    # Comando para criar o executável
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--name", "WallEngenhariaApp",
        "--distpath", dist_dir,
        "--workpath", os.path.join(base_dir, "build"),
        "--specpath", os.path.join(base_dir, "build"),
        "--clean",
        main_script
    ]
    
    # Executar o comando
    print("Criando executável...")
    subprocess.check_call(cmd)
    
    return True

def create_portable_package():
    """Cria um pacote portátil do aplicativo."""
    # Diretório base do projeto
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    # Diretório de saída para o pacote portátil
    portable_dir = os.path.join(base_dir, "dist", "WallEngenhariaApp_Portable")
    if not os.path.exists(portable_dir):
        os.makedirs(portable_dir)
    
    # Copiar executável
    exe_path = os.path.join(base_dir, "dist", "WallEngenhariaApp.exe")
    if os.path.exists(exe_path):
        shutil.copy(exe_path, portable_dir)
    
    # Copiar recursos
    resources_dir = os.path.join(base_dir, "resources")
    if os.path.exists(resources_dir):
        dest_resources = os.path.join(portable_dir, "resources")
        if os.path.exists(dest_resources):
            shutil.rmtree(dest_resources)
        shutil.copytree(resources_dir, dest_resources)
    
    # Copiar documentação
    docs_dir = os.path.join(base_dir, "docs")
    if os.path.exists(docs_dir):
        dest_docs = os.path.join(portable_dir, "docs")
        if os.path.exists(dest_docs):
            shutil.rmtree(dest_docs)
        shutil.copytree(docs_dir, dest_docs)
    
    # Criar arquivo README.txt
    with open(os.path.join(portable_dir, "README.txt"), "w", encoding="utf-8") as f:
        f.write("Aplicativo de Integração Excel-Word para Wall Engenharia\n")
        f.write("=======================================================\n\n")
        f.write("Para iniciar o aplicativo, execute o arquivo WallEngenhariaApp.exe\n\n")
        f.write("Para mais informações, consulte a documentação na pasta docs/\n")
    
    # Criar arquivo ZIP
    import zipfile
    zip_path = os.path.join(base_dir, "dist", "WallEngenhariaApp_Portable.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(portable_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, portable_dir)
                zipf.write(file_path, arcname)
    
    print(f"Pacote portátil criado: {zip_path}")
    return True

def main():
    """Função principal."""
    # Verificar sistema operacional
    if platform.system() != "Windows":
        print("Aviso: Este script foi projetado para Windows. O empacotamento em outros sistemas pode não funcionar corretamente.")
    
    # Verificar requisitos
    check_requirements()
    
    # Criar executável
    if create_executable():
        print("Executável criado com sucesso!")
        
        # Criar pacote portátil
        if create_portable_package():
            print("Pacote portátil criado com sucesso!")
    else:
        print("Erro ao criar executável.")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
