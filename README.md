# PlanYazo - Sistema de Transferência de Dados

Sistema gráfico para transferência de dados entre planilhas Excel, desenvolvido para gerenciar eventos e palestras.

## 📋 Descrição

O PlanYazo é uma aplicação desktop com interface gráfica moderna que permite transferir dados da planilha `Plan.xlsx` para `Yazo.xlsx`, especificamente para gerenciar eventos como palestras, workshops e painéis.

## ✨ Funcionalidades

### 🎯 Transferência de Dados
- **Palestras**: Transfere dados da aba PALESTRA
- **Workshops**: Transfere dados da aba WORKSHOP  
- **Painéis**: Transfere dados da aba PAINEL
- **Apagar Dados**: Remove todos os registros da aba Palestras2k25

### 🔍 Características Técnicas
- Interface gráfica moderna e responsiva
- Verificação automática de arquivos
- Processamento em threads para não travar a interface
- Log detalhado de todas as operações
- Detecção de células com fundo verde (RGB: 0,255,0)
- Confirmações de segurança para operações críticas

## 🚀 Instalação

### Pré-requisitos
- Python 3.7 ou superior
- Windows 10/11 (testado)

### Passos de Instalação

1. **Clone ou baixe o projeto**
2. **Instale as dependências:**
```bash
pip install -r requirements.txt
```

3. **Execute a aplicação:**
```bash
python Front.py
```

## 📦 Criando Executável

### Método Automático (Recomendado)
1. **Execute o script batch:**
```bash
criar_executavel.bat
```

### Método Manual
1. **Execute o script Python:**
```bash
python build_exe.py
```

### Resultado
- O executável será criado na pasta `PlanYazo_Executavel/`
- Inclui: `PlanYazo.exe`, `Plan.xlsx`, `Yazo.xlsx` e `README_Executavel.txt`
- Para usar: Dê duplo clique em `PlanYazo.exe`

## 📁 Estrutura de Arquivos

```
HTPlanYazo/
├── Front.py                 # Interface gráfica principal
├── requirements.txt         # Dependências do projeto
├── build_exe.py            # Script para criar executável
├── criar_executavel.bat    # Script batch para Windows
├── PlanYazo.spec           # Configuração do PyInstaller
├── Plan.xlsx               # Planilha fonte de dados
├── Yazo.xlsx               # Planilha de destino
├── README.md               # Este arquivo
└── Especificacoes.md       # Especificações detalhadas

PlanYazo_Executavel/        # Pasta gerada automaticamente
├── PlanYazo.exe            # Executável da aplicação
├── Plan.xlsx               # Planilha fonte (cópia)
├── Yazo.xlsx               # Planilha destino (cópia)
└── README_Executavel.txt   # Instruções de uso
```

## 🎮 Como Usar

### 1. Preparação dos Arquivos
- Coloque `Plan.xlsx` e `Yazo.xlsx` na mesma pasta do programa
- Certifique-se que os arquivos não estão abertos no Excel

### 2. Execução
- Execute `python Front.py`
- A interface verificará automaticamente se os arquivos estão presentes

### 3. Transferência de Dados

#### Para Palestras, Workshops ou Painéis:
1. Clique no botão correspondente (🎤 Palestra, 🔧 Workshop, 👥 Painel)
2. O sistema irá:
   - Procurar linhas com fundo verde na coluna A da aba correspondente
   - Extrair os dados mapeados
   - Transferir para a aba "Palestras2k25" do Yazo.xlsx

#### Para Apagar Dados:
1. Clique em "🗑️ Apagar Dados"
2. Confirme a operação (apaga todos os registros mantendo o cabeçalho)

### 4. Monitoramento
- Use a área de log para acompanhar as operações
- O botão "🔄 Atualizar" verifica novamente o status dos arquivos

## 📊 Mapeamento de Dados

### Estrutura de Transferência
| Plan.xlsx (Coluna) | Yazo.xlsx (Campo) | Descrição |
|-------------------|-------------------|-----------|
| A | Nome dos Palestrantes | Nome do palestrante |
| N | Local | Local do evento |
| F | Título* | Título da apresentação |
| G | Descrição | Descrição detalhada |
| K | Data (YYYY/MM/DD)* | Data do evento |
| L | Horário de início | Hora de início |
| M | Horário de término | Hora de término |

### Critérios de Seleção
- **Apenas linhas com fundo verde** na coluna A são processadas
- Cor específica: RGB(0,255,0) / #00FF00
- Linhas sem nome e título são ignoradas

## 🔧 Configurações

### Interface
- Tamanho da janela: 1200x800 pixels
- Tema: Clam (moderno)
- Cores: Esquema azul/cinza profissional

### Log
- Timestamp automático em todas as mensagens
- Scroll automático para última mensagem
- Botão para limpar histórico

## ⚠️ Observações Importantes

### Segurança
- Confirmação obrigatória para operações críticas
- Aviso quando arquivos estão abertos no Excel
- Rastreamento de opções já utilizadas na sessão

### Compatibilidade
- Testado com Excel 2016+
- Requer arquivos .xlsx (não .xls)
- Funciona apenas com Windows

### Performance
- Processamento em threads para não travar a interface
- Verificação eficiente de cores de células
- Log detalhado para debugging

## 🐛 Solução de Problemas

### Arquivo não encontrado
- Verifique se `Plan.xlsx` e `Yazo.xlsx` estão na pasta correta
- Use o botão "🔄 Atualizar" para verificar novamente

### Erro de permissão
- Feche os arquivos Excel antes de executar
- Verifique se não há outros programas usando os arquivos

### Nenhuma linha encontrada
- Confirme que existem células com fundo verde na coluna A
- Verifique se está na aba correta (PALESTRA, WORKSHOP, PAINEL)

## 📝 Log de Alterações

### Versão Atual
- Interface gráfica moderna
- Processamento em threads
- Verificação automática de arquivos
- Log detalhado de operações
- Confirmações de segurança

### Versões Anteriores
- Versão console (transfer_data.py)
- Processamento sequencial
- Interface básica

## 🤝 Contribuição

Para contribuir com o projeto:
1. Faça um fork do repositório
2. Crie uma branch para sua feature
3. Commit suas mudanças
4. Abra um Pull Request

## 📄 Licença

Este projeto é de uso interno para gerenciamento de eventos.

---

**Desenvolvido para Hacktown 2025** 🚀


            

            

            
