# Objetivo:
Alterar o script para operar nos seguintes parâmetros:
	- Menu Inicial:
        - Opções de importação para Yazo:
            1 - Palestra
            2 - Workshop
            3 - Painel
            4 - Fechar
            5 - Apagar tudo

        Lógica (Na parte da Yazo, os registros devem ser contidos na aba "2025", abaixo do cabeçalho): 
            ## PARA TODAS AS OPÇÕES, CAPTAR SOMENTE LINHAS QUE A COLUNA 1 ESTÁ VERDE (#00FF00 ou 0,255,0)

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
                - Apagar todas as celulas abaixo do cabeçalho da Planilha Yazo, aba 2025

            
            Detalhe importante:
            Vamos supor que eu inicie a aplicação, escolha a opção 1 e a mesma registre 25 linhas abaixo do cabeçalho:
                - Após confirmar o registro, o menu de interação deve voltar com as opções novamente, e se eu escolher a opção 2, por exemplo, ele deve fazer os novos registros abaixo dos 25 que foram registrados pela opção 1
                - Fazer o mesmo para todas as opções
                - Somente fechar a aplicação quando eu optar pela opção 4
                - Se na mesma interação, eu selecionar 2x a mesma opção, alertar o usuario se deseja prosseguir e pedir confirmação
                - Se selecionado a op 5, alertar o usuario se deseja prosseguir e pedir confirmação deixando evidente que os registros serão apagados.


            

            

            
