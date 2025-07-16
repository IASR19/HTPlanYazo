# Programa de Transferência de Dados Excel

Este programa transfere dados da planilha `Plan.xlsx` para a planilha `Yazo.xlsx`.

## Instalação

1. Certifique-se de ter Python 3.7+ instalado
2. Instale as dependências:
```bash
pip install -r requirements.txt
```

## Como usar

1. Coloque os arquivos `Plan.xlsx` e `Yazo.xlsx` na mesma pasta do programa
2. Execute o programa:
```bash
python transfer_data.py
```

# Objetivo:
Alterar o script para operar nos seguintes parâmetros:
	- Menu Inicial:
        - Opções de importação para Yazo:
            1 - Palestra
            2 - Workshop
            3 - Painel
            4 - Fechar
            5 - Apagar tudo

        Lógica (Na parte da Yazo, os registros devem ser contidos na aba "Palestras2k25", abaixo do cabeçalho): 
            ## PARA TODAS AS OPÇÕES, CAPTAR SOMENTE LINHAS QUE ESTÃO COM TODAS CELULAS DA LINHA COM PREENCHIMENTO DE FUNDO EM VERDE

            - Se selecionado a op 1, a planilha Yazo deverá obter dados da Plan (aba PALESTRA) nos seguintes critérios: Plan (Coluna) -> Yazo (Coluna)
                - Nome (A) -> Nome do palestrante (N)
                - Local (N) -> Local (B)
                - Titulo (F) -> Titulo (E)
                - Descritivo (G) -> Descrição (H)
                - Data (K) -> Data (K)
                - Hora de inicio (L) -> Hora de inicio (L)
                - Hora de fim (M) -> Hora de fim (M)

            - Se selecionado a op 2, a planilha Yazo deverá obter dados da Plan (aba WORKSHOP) nos seguintes critérios: Plan (Coluna) -> Yazo (Coluna)
                - Nome (A) -> Nome (N)
                - Local (N) -> Local (B)
                - Titulo (F) -> Titulo (E)
                - Descritivo (G) -> Descrição (H)
                - Data (K) -> Data (K)
                - Hora de inicio (L) -> Hora de inicio (L)
                - Hora de fim (M) -> Hora de fim (M)

            - Se selecionado a op 3, a planilha Yazo deverá obter dados da Plan (aba PAINEL) nos seguintes critérios: Plan (Coluna) -> Yazo (Coluna)
                - Nome (A) -> Nome (N)
                - Local (N) -> Local (B)
                - Titulo (F) -> Titulo (E)
                - Descritivo (G) -> Descrição (H)
                - Data (K) -> Data (K)
                - Hora de inicio (L) -> Hora de inicio (L)
                - Hora de fim (M) -> Hora de fim (M)

            - Se selecionado a op 4 
                - Fechar a aplicação
            
            - Se selecionado a op 5
                - Apagar todas as celulas abaixo do cabeçalho da Planilha Yazo, aba Palestras2k25

            
            Detalhe importante:
            Vamos supor que eu inicie a aplicação, escolha a opção 1 e a mesma registre 25 linhas abaixo do cabeçalho:
                - Após confirmar o registro, o menu de interação deve voltar com as opções novamente, e se eu escolher a opção 2, por exemplo, ele deve fazer os novos registros abaixo dos 25 que foram registrados pela opção 1
                - Fazer o mesmo para todas as opções
                - Somente fechar a aplicação quando eu optar pela opção 4
                - Se na mesma interação, eu selecionar 2x a mesma opção, alertar o usuario se deseja prosseguir e pedir confirmação
                - Se selecionado a op 5, alertar o usuario se deseja prosseguir e pedir confirmação deixando evidente que os registros serão apagados.


            

            

            
