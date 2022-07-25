import os
import shutil
import time
import pdfkit 
import unidecode 
import pandas as pd
import configLeticia
from datetime import date, datetime, timedelta
from chave import chave_api


from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementClickInterceptedException
from anticaptchaofficial.imagecaptcha import *






class Bot():

    def __init__(self):
        self.data = datetime.now()
        self.reset_ambiente()
        self.driver = self.get_driver()               
        self.data_pasta = self.data.strftime('%d-%m-%Y')
        print(self.data_pasta)
       
        self.run()


    def get_driver(self):
        chrome_options = Options()
        #chrome_options.add_argument('--headless')
        chrome_options.add_experimental_option("prefs", {"download.default_directory": configLeticia.DIRETORIO_ARQUIVOS_TEMP})
        s=Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=s,options=chrome_options)
        driver.maximize_window()
        return driver

    def cria_diretorio(self, caminho):
        if not os.path.isdir(caminho):
            os.mkdir(caminho)

    def reset_ambiente(self):
        list_caminhos = [
            configLeticia.DIRETORIO_ARQUIVOS_TEMP
                   ]
        for caminho in list_caminhos:
            if os.path.isdir(caminho):
                shutil.rmtree(caminho)
            self.cria_diretorio(caminho)

    def login_smartsheet(self, usuario, senha):
        self.driver.get(configLeticia.URL_SMARTSHEET)
        #Insere e-mail
        self.retorna_elemento('ID', 'loginEmail').send_keys(usuario)
        #clica em entrar
        self.retorna_elemento('ID', 'formControl').click()
        #insere senha
        self.retorna_elemento('ID', 'loginPassword').send_keys(senha)
        #clica em entrar
        self.retorna_elemento('ID', 'formControl').click()

        #tempo necessário pois o carregamento do site demora
        time.sleep(1)

    
    def aguarda_download(self):
        seconds = 1
        time.sleep(seconds)
        dl_wait = True
        while dl_wait and seconds < 60:
            time.sleep(2)
            dl_wait = False
            for fname in os.listdir(configLeticia.DIRETORIO_ARQUIVOS_TEMP):
                if fname.endswith('.crdownload'):
                    dl_wait = True
            seconds += 1
        return seconds

    def renomar_arquivo(self, original, alterar):
        #self.convert_to_parquet(original)
        self.aguarda_download()
        os.rename(original, alterar)

   
    def cria_diretorio(self, caminho):
        if not os.path.isdir(caminho):
            os.mkdir(caminho)

    def retorna_elemento(self, funcao, path):
        self.aguardar_elemento(funcao, path)
        return self.driver.find_element(getattr(By,funcao), path)

    def aguardar_elemento(self, funcao, path):
        WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((getattr(By,funcao), path)))

    def remover_notificacao(self):
        time.sleep(3)
        elements = self.driver.find_elements(By.TAG_NAME, 'tecnofit-micro-notification')
        for e in elements:
            self.driver.execute_script("""
            var element = arguments[0];
            element.parentNode.removeChild(element);
            """, e)


    def extrair_smartsheet(self):
        print('Extraindo planilha smartsheet...')

        time.sleep(10)
        #acessa Planilha:
        self.driver.get('https://app.smartsheet.com/sheets/qxXHMXfCQf3JRHhRGXfFwxqwVRHmj6v6cPJh9Q31?view=card&cardLevel=0&cardViewByColumnId=6596899369183108')

        try:
            #fecha modal
            self.retorna_elemento('XPATH', 'scButton scTextButton tertiary').click()
        except:
            print('Não apareceu Modal!')
        #clica em Arquivo
        self.retorna_elemento('XPATH', '//*[@id="mnb-1"]').click()
        #clica em exportar
        self.retorna_elemento('XPATH', "//tr[@data-client-id='10621']").click()
        #clica em exportar para excel
        self.retorna_elemento('XPATH', "//tr[@data-client-id='10604']").click()
        
        #Aguarda Download
        print('Aguardando Download...')
        self.aguarda_download()


        #renomeia e move arquivo
        print('Renomeando Arquivo...')
        caminho_inicio = configLeticia.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'LICITAÇÕES.xlsx'
        caminho_fim = configLeticia.DIRETORIO_SMARTSHEET + '\\' + self.data_pasta + '\\' + 'LICITACOES.xlsx' 
        self.renomar_arquivo(caminho_inicio, caminho_fim)

    
    def extracao_comprasnet(self):
        print('Extraindo relatório comprasNEt...')

        #acessando planilha smartsheet
        caminho_fim = configLeticia.DIRETORIO_SMARTSHEET + '\\' + self.data_pasta + '\\' + 'LICITACOES.xlsx'
        arq = pd.read_excel(caminho_fim)

        #separando informaões necessárias
        selecao = arq[['Cliente', 'Código', 'Número da licitação', 'Local', 'Identificação' ]]
        filtro = selecao[selecao['Local']=='comprasnet']

        #Início do laço
        for codUASG,numPreg, cliente, identificacao in zip(filtro['Código'], filtro['Número da licitação'], filtro['Cliente'], filtro['Identificação']):
            #acessa area do site de consultas
            self.driver.get('http://comprasnet.gov.br/livre/Pregao/lista_pregao_filtro.asp?Opc=2')
            time.sleep(1)

            #Seleciona Situação = Todos
            self.retorna_elemento('XPATH', "//option[@value='5']").click()
            time.sleep(1)
            #Insere Cod. UASG
            print(codUASG)
            self.retorna_elemento('ID', 'co_uasg').send_keys(codUASG)
            time.sleep(1)
            #Insere Numero Pregao
            print(numPreg)
            self.retorna_elemento('ID', 'numprp').send_keys(numPreg)
            time.sleep(1)

            #Clica em OK para buscar as informações
            self.retorna_elemento('ID', 'ok').click()
            time.sleep(1)
            
            try:
                self.retorna_elemento('XPATH', "//a[@href='#']")
                #clica no pregao informado
                if(len(self.driver.find_elements(By.XPATH, '/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr'))==2):
                    self.retorna_elemento('XPATH', "//a[@href='#']").click()
                else :
                    print('Faltou o código do pregão!')
                    continue
            except TimeoutException:
                print('Não existem pregões no momento')
                continue


            #resolve captcha
            self.quebra_captcha(cliente, identificacao)

      

    def quebra_captcha(self, cliente, identificacao):
        print('Quebrando Captcha...')

        solver = imagecaptcha()
        solver.set_verbose(1)
        solver.set_key(chave_api)

        # Specify softId to earn 10% commission with your app.
        # Get your softId here: https://anti-captcha.com/clients/tools/devcenter
        solver.set_soft_id(0)

        #obtem a imagem do captcha
        element = self.retorna_elemento('XPATH', '//*[@id="form1"]/table/tbody/tr[1]/td/table/tbody/tr/td[2]/span/img')
        element.screenshot(r'C:\Programação\leticia\src\captcha.png')

        captcha_text = solver.solve_and_return_solution("C:\Programação\leticia\src\captcha.png")
        if captcha_text != 0:
            print ("captcha text "+captcha_text)
        else:
            print ("task finished with error "+solver.error_code)

        #insere texto captcha na caixa
        self.retorna_elemento('ID', 'idLetra').send_keys(captcha_text)

        
        #clica em confirmar
        element_to_hover_over = self.retorna_elemento('ID', "idSubmit")
        hover = ActionChains(self.driver).move_to_element(element_to_hover_over)
        hover.perform()
        time.sleep(1)
        self.retorna_elemento('ID', 'idSubmit').click()
        time.sleep(5)

        #caso o código inserido seja errado
        try:
            self.retorna_elemento('XPATH', "//input[@name='chat']")
            consultas = self.driver.find_elements(By.XPATH, '/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[6]/td/table/tbody/tr/td/a[text()]')
            contador = 1
            for i in consultas:
                if(i.text == 'Realizar julgamento'):
                    #clicando na situação do pregao
                    i.click()
                    print(contador)
                    
                    #mudando para a jalena que foi aberta
                    janelas = self.driver.window_handles
                    self.driver.switch_to.window(janelas[1])

                    #transformando a página em html        
                    html = self.driver.page_source 

                    #transformando as tabelas que a página contem em DataFrame
                    pagina=pd.read_html(html)
                    df = pd.DataFrame(pagina[0])

                    #Salvando em arquivo Excel
                    caminhoFim = (configLeticia.DIRETORIO_ARQUIVOS + '\\'+ self.data_pasta  +'\\' +'relatorio_comprasNet' + '_'+ cliente + '_' + identificacao + '_consulta' + str(contador) + '.xlsx')
                    df.to_excel(caminhoFim)
                    time.sleep( 2 )     

                    #fechando janela que foi aberta
                    self.driver.close()
                    self.driver.switch_to.window(janelas[0])
                    contador += 1
        except TimeoutException:
            print('captcha inserido de forma errada ou o site não aceitou a solicitação. Tentando novamente...')
            self.quebra_captcha(cliente, identificacao)    
            
        

            

    def run(self):
        for credencial in configLeticia.CREDENCIAIS:
            self.login_smartsheet(credencial['usuario'], credencial['senha'])
            self.cria_diretorio(configLeticia.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta) 
            self.cria_diretorio(configLeticia.DIRETORIO_SMARTSHEET + '\\' + self.data_pasta)           
            self.extrair_smartsheet()  
            self.extracao_comprasnet()              
        self.driver.close()
        self.driver.quit()

if __name__ == '__main__':
    bot = Bot()
