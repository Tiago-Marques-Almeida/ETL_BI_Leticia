# ETL_BI_Leticia
Projeto que envolve a criação de uma base de dados, análise de dados e criação de painel em Power BI

A criação da Base de dados é feita através de um robô criado em Python. Inicialmente ele extrai uma planilha do site Smartsheet. Depois ele lê os dados dessa planilha para conseguir fazer a consulta no site compras net (site de licitações do governo). Esse site requer a quebra de Captcha, então coloquei uma função que o faz. Por fim utilizo a biblioteca Pandas para deixar o arquivo extraído no formato para ser lido no Power BI.
