# Aplicativo-Excel-Word-Wall-Engenharia
App para integrar as informações do excel, ao Word
# README - Aplicativo de Integração Excel-Word para Wall Engenharia

## Descrição

Este aplicativo foi desenvolvido para a Wall Engenharia com o objetivo de automatizar o processo de criação de propostas técnicas e comerciais a partir de planilhas de orçamento. O sistema integra dados de planilhas Excel diretamente em modelos de documentos Word, permitindo a geração rápida e padronizada de propostas comerciais.

## Estrutura do Projeto

```
wall_engenharia_app/
│
├── src/                      # Código-fonte
│   ├── main.py               # Ponto de entrada da aplicação
│   ├── excel_reader.py       # Módulo de leitura de Excel
│   ├── word_writer.py        # Módulo de manipulação de Word
│   ├── ui_manager.py         # Módulo de interface de usuário
│   └── config_manager.py     # Módulo de configuração
│
├── resources/                # Recursos estáticos
│   ├── templates/            # Modelos de proposta
│   └── config/               # Arquivos de configuração
│
├── tests/                    # Testes automatizados
│   └── test_integration.py   # Testes unitários e de integração
│
├── docs/                     # Documentação
│   ├── manual_usuario.md     # Manual do usuário
│   └── manual_tecnico.md     # Documentação técnica
│
└── scripts/                  # Scripts de utilidade
    └── package.py            # Script para empacotamento do aplicativo
```

## Requisitos

### Para Desenvolvimento
- Python 3.8 ou superior
- Bibliotecas: openpyxl, python-docx, pandas, tkinter
- PyInstaller (para empacotamento)

### Para Usuários Finais
- Windows 10 ou superior
- Microsoft Office (Excel e Word) 2016 ou superior
- Mínimo de 4 GB de RAM recomendado
- Mínimo de 100 MB de espaço em disco

## Instalação para Desenvolvimento

1. Clone o repositório:
```
git clone https://github.com/wallengenharia/excel-word-integration.git
cd excel-word-integration
```

2. Instale as dependências:
```
pip install -r requirements.txt
```

3. Execute o aplicativo:
```
python src/main.py
```

## Empacotamento para Distribuição

Para criar um executável para Windows:

1. Execute o script de empacotamento:
```
python scripts/package.py
```

2. Os arquivos de distribuição serão gerados no diretório `dist/`:
   - `WallEngenhariaApp.exe`: Executável standalone
   - `WallEngenhariaApp_Portable.zip`: Versão portátil do aplicativo

## Documentação

- **Manual do Usuário**: Disponível em `docs/manual_usuario.md`
- **Documentação Técnica**: Disponível em `docs/manual_tecnico.md`

## Testes

Para executar os testes automatizados:

```
python -m unittest tests/test_integration.py
```

## Funcionalidades Principais

- Extração de dados de planilhas Excel de orçamento
- Preenchimento automático de modelos de proposta Word
- Interface gráfica intuitiva para seleção de arquivos e edição de dados
- Personalização de modelos através de placeholders
- Gerenciamento de arquivos recentes
- Visualização prévia dos dados antes da geração da proposta

## Contato

Para suporte ou mais informações, entre em contato com:

- **E-mail**: suporte@wallengenharia.com.br
- **Telefone**: (XX) XXXX-XXXX

## Licença

Este software é propriedade da Wall Engenharia e seu uso é restrito aos termos estabelecidos no contrato de licença.
