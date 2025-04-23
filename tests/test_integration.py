#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script de teste para o aplicativo Wall Engenharia.
Testa as funcionalidades básicas dos módulos de integração Excel-Word.
"""

import os
import sys
import unittest
import tempfile
import shutil

# Adicionar diretório pai ao path para importar módulos
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Importar módulos do aplicativo
from src.excel_reader import ExcelReader, extract_data_from_excel
from src.word_writer import WordWriter, generate_proposal
from src.config_manager import get_config_manager

class TestExcelReader(unittest.TestCase):
    """Testes para o módulo de leitura de Excel."""
    
    def setUp(self):
        """Configuração inicial para os testes."""
        # Caminho para o arquivo Excel de exemplo
        self.excel_file = "/path/to/your/excel/file.xlsx"  # Substitua pelo caminho real do arquivo
        
        # Verificar se o arquivo existe
        self.assertTrue(os.path.exists(self.excel_file), f"Arquivo de teste não encontrado: {self.excel_file}")
    
    def test_excel_reader_init(self):
        """Testa a inicialização do leitor de Excel."""
        reader = ExcelReader(self.excel_file)
        self.assertIsNotNone(reader)
        self.assertIsNotNone(reader.workbook)
        self.assertGreater(len(reader.sheet_names), 0)
        reader.close()
    
    def test_get_sheet_names(self):
        """Testa a obtenção dos nomes das planilhas."""
        reader = ExcelReader(self.excel_file)
        sheet_names = reader.get_sheet_names()
        self.assertIsInstance(sheet_names, list)
        self.assertGreater(len(sheet_names), 0)
        reader.close()
    
    def test_select_sheet(self):
        """Testa a seleção de uma planilha."""
        reader = ExcelReader(self.excel_file)
        result = reader.select_sheet(reader.sheet_names[0])
        self.assertTrue(result)
        self.assertIsNotNone(reader.current_sheet)
        reader.close()
    
    def test_extract_data_for_proposal(self):
        """Testa a extração de dados para a proposta."""
        reader = ExcelReader(self.excel_file)
        reader.select_sheet(reader.sheet_names[0])
        data = reader.extract_data_for_proposal()
        self.assertIsInstance(data, dict)
        self.assertIn("nome_cliente", data)
        reader.close()
    
    def test_extract_data_from_excel_function(self):
        """Testa a função auxiliar para extração de dados."""
        data = extract_data_from_excel(self.excel_file, None)
        self.assertIsInstance(data, dict)
        self.assertIn("nome_cliente", data)


class TestWordWriter(unittest.TestCase):
    """Testes para o módulo de manipulação de Word."""
    
    def setUp(self):
        """Configuração inicial para os testes."""
        # Caminho para o arquivo Word de exemplo
        self.word_file = "/path/to/your/word/template.docx"  # Substitua pelo caminho real do arquivo
        
        # Verificar se o arquivo existe
        self.assertTrue(os.path.exists(self.word_file), f"Arquivo de teste não encontrado: {self.word_file}")
        
        # Criar diretório temporário para arquivos de saída
        self.temp_dir = tempfile.mkdtemp()
        self.output_file = os.path.join(self.temp_dir, "proposta_teste.docx")
    
    def tearDown(self):
        """Limpeza após os testes."""
        # Remover diretório temporário
        shutil.rmtree(self.temp_dir)
    
    def test_word_writer_init(self):
        """Testa a inicialização do escritor de Word."""
        writer = WordWriter(self.word_file)
        self.assertIsNotNone(writer)
        self.assertIsNotNone(writer.document)
    
    def test_fill_proposal_with_data(self):
        """Testa o preenchimento da proposta com dados."""
        writer = WordWriter(self.word_file)
        
        # Dados de teste
        test_data = {
            "nome_cliente": "Empresa Teste",
            "nome_contato": "João Silva",
            "email": "joao@teste.com",
            "telefone": "(11) 98765-4321",
            "escopo": "Serviço de teste",
            "prazo": "30 dias",
            "custo": 50000.00,
            "garantias": "12 meses",
            "seguro": "Incluído",
            "nao_inclusos": ["Item 1", "Item 2"]
        }
        
        # Preencher proposta
        result = writer.fill_proposal_with_data(test_data)
        self.assertTrue(result)
        
        # Salvar documento
        save_result = writer.save_document(self.output_file)
        self.assertTrue(save_result)
        self.assertTrue(os.path.exists(self.output_file))
    
    def test_generate_proposal_function(self):
        """Testa a função auxiliar para geração de proposta."""
        # Dados de teste
        test_data = {
            "nome_cliente": "Empresa Teste",
            "nome_contato": "João Silva",
            "email": "joao@teste.com",
            "telefone": "(11) 98765-4321",
            "escopo": "Serviço de teste",
            "prazo": "30 dias",
            "custo": 50000.00,
            "garantias": "12 meses",
            "seguro": "Incluído",
            "nao_inclusos": ["Item 1", "Item 2"]
        }
        
        # Gerar proposta
        result = generate_proposal(self.word_file, self.output_file, test_data)
        self.assertTrue(result)
        self.assertTrue(os.path.exists(self.output_file))


class TestIntegration(unittest.TestCase):
    """Testes de integração entre os módulos."""
    
    def setUp(self):
        """Configuração inicial para os testes."""
        # Caminhos para os arquivos de teste
        self.excel_file = "/path/to/your/excel/file.xlsx"  # Substitua pelo caminho real do arquivo
        self.word_file = "/path/to/your/word/template.docx"  # Substitua pelo caminho real do arquivo
        
        # Verificar se os arquivos existem
        self.assertTrue(os.path.exists(self.excel_file), f"Arquivo Excel de teste não encontrado: {self.excel_file}")
        self.assertTrue(os.path.exists(self.word_file), f"Arquivo Word de teste não encontrado: {self.word_file}")
        
        # Criar diretório temporário para arquivos de saída
        self.temp_dir = tempfile.mkdtemp()
        self.output_file = os.path.join(self.temp_dir, "proposta_integrada.docx")
    
    def tearDown(self):
        """Limpeza após os testes."""
        # Remover diretório temporário
        shutil.rmtree(self.temp_dir)
    
    def test_excel_to_word_integration(self):
        """Testa a integração completa de Excel para Word."""
        # Extrair dados do Excel
        data = extract_data_from_excel(self.excel_file)
        self.assertIsInstance(data, dict)
        
        # Complementar dados que podem estar faltando
        if not data.get("nome_cliente"):
            data["nome_cliente"] = "Empresa Teste"
        
        if not data.get("nome_contato"):
            data["nome_contato"] = "João Silva"
        
        if not data.get("email"):
            data["email"] = "contato@exemplo.com"
        
        if not data.get("telefone"):
            data["telefone"] = "(11) 98765-4321"
        
        if not data.get("escopo"):
            data["escopo"] = "Serviço de teste"
        
        if not data.get("prazo"):
            data["prazo"] = "30 dias"
        
        if not data.get("custo") or data["custo"] is None:
            data["custo"] = 50000.00
        
        # Gerar proposta com os dados extraídos
        result = generate_proposal(self.word_file, self.output_file, data)
        self.assertTrue(result)
        self.assertTrue(os.path.exists(self.output_file))
        
        # Verificar tamanho do arquivo gerado
        file_size = os.path.getsize(self.output_file)
        self.assertGreater(file_size, 10000)  # Arquivo deve ter um tamanho razoável


if __name__ == "__main__":
    unittest.main()
