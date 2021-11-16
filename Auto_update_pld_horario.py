# -*- coding: utf-8 -*-
"""
Created on Fri Nov  5 07:58:14 2021

@author: 2013920
    Lucas Leite Tavares - Estagiário de Inteligência de Mercado
    
    INFORMAÇOES IMPORTANTES:
        
    (I)  - Esse Script trabalha com web_driver, certifique-se que a pasta do 
    chrome_driver.exe esteja atualizada.
    (II) - Preencha os campos de entreda corretamente, pois eles são utilizados
    como localização dos diretorios.
    
"""

#%% Inicio
# Script downloads CCEE WEBSITE
    #Implementados:
        #PLD diario direto do site da CCEE
# Lucas Leite Tavares
# Version 0.1

import os
import os.path
import time
import cx_Oracle
import locale
from selenium import webdriver
import win32com.client as win32
from datetime import datetime
from datetime import timedelta
locale.setlocale(locale.LC_ALL, 'pt')
from selenium.webdriver.chrome.options import Options



#%% Credenciais de entrada:
    
data_hoje = (datetime.today().strftime('%d_%m_%Y'))
data_ontem= ((datetime.today() - timedelta(days=1)).strftime('%d_%m_%Y'))
data_e_hora = datetime.today().strftime('%d/%m/%Y %H:%M')

#print('É necessário informar suas credenciais para se comunicar...um dia com o API')
#usuario = input("Usuário CCEE:")
#senha = getpass.getpass("Senha CCEE:")

#matricula = input("Matrícula CPFL:")
matricula = 2013920

#Insere data e transforma em date.time
#201date = input("Data(dd/mm/aaaa)):")
#date = datetime.datetime.strptime(date, "%d/%m/%Y")

#Define o caminho do chromedriver. Precisa baixar e trocar o arquivo se for a 1ª vez usando selenium
path_nav = rf"C:\Users\2013920\AppData\Local\Google\Chrome\User Data/chromedriver.exe"

#Define diretório de download do usuário
dirdown = rf"C:/Users/{2013920}/Downloads/"
#%% Utilização do navegador:

# Opções e abertura do navegador
options = Options()
options.add_argument("start-maximized")
options.binary_location = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
driver = webdriver.Chrome(path_nav, options = options)
#Fonte da Pasta de Trabalho: https://tableaupub.ccee.org.br/t/CCEE/views/PreoHorriodoDia/PreoHorriodoDia?:embed=y&:showVizHome=no&:host_url=https%3A%2F%2Ftableaupub.ccee.org.br%2F&:embed_code_version=3&:tabs=yes&:toolbar=yes&:iid=1&:isGuestRedirectFromVizportal=y&:display_count=no&:origin=viz_share_link&:showShareOptions=false&:alerts=no&:refresh=yes&:display_spinner=no&:loadOrderID=4

driver.get("https://tableaupub.ccee.org.br/t/CCEE/views/PreoHorriodoDia/Download?%3Aembed=y&%3AshowVizHome=no&%3Ahost_url=https%3A%2F%2Ftableaupub.ccee.org.br%2F&%3Aembed_code_version=3&%3Atabs=yes&%3Atoolbar=yes&%3Aiid=1&%3AisGuestRedirectFromVizportal=y&%3Adisplay_count=no&%3Aorigin=viz_share_link&%3AshowShareOptions=false&%3Aalerts=no&%3Arefresh=yes&%3Adisplay_spinner=no&%3AloadOrderID=4")
#driver.get("https://tableaupub.ccee.org.br/vizql/t/CCEE/w/PreoHorriodoDia/v/Download/vud/sessions/D0D3EA8539F44A548D829FA8546FC680-1:1/views/4615802733734373727_8080701578547368505?csv=true&summary=true%22")
time.sleep(10)

#%% Sequência automatizada de baixar arquivo:


# Seleciona Ontem
button = driver.find_element_by_xpath("//*[@id='FI_federated.1y14jvx0g56vcm1cdv0uy0517dzt,none:FLAG:nk4615802733734373727_8080701578547368505_1']")
button.click()
time.sleep(5)

# Tira o check de Ontem
button = driver.find_element_by_xpath("//*[@id='FI_federated.1y14jvx0g56vcm1cdv0uy0517dzt,none:FLAG:nk4615802733734373727_8080701578547368505_1']/div[2]/input")
button.click()
time.sleep(5)

# Clica em aplicar
button = driver.find_element_by_xpath("//*[@id='tableau_base_widget_LegacyCategoricalQuickFilter_1']/div/div[3]/div[3]/button[2]")
button.click()
time.sleep(10)

# Clica em baixar
button = driver.find_element_by_xpath("//*[@id='toolbar-container']/div[1]/div[2]/div[1]")
button.click()
time.sleep(5)

# Seleciona o tipo de arquivo
button = driver.find_element_by_xpath("//*[@id='DownloadDialog-Dialog-Body-Id']/div/div[2]")
button.click()
time.sleep(5)
#%% Trocar as abas do navegador
all_windows = window_after = driver.window_handles

window_before = all_windows[0]
window_after = all_windows[1]


driver.switch_to.window(window_after)
#driver = driver.switch_to_window(driver2)

# Click no link de download:
button = driver.find_element_by_xpath("//*[@id='tabContent-panel-summary']/div[1]/div[2]/a")
button.click()
time.sleep(5)

driver.quit()


#%% Criando arquivo CSV de auxilio:

save_path = r'C:\Users\2013920\Downloads'
name_of_file = "Tabela_data1"
completeName = os.path.join(save_path, name_of_file+".csv")         
file1 = open(completeName, "w")
file1.close()

#%%  Endereços de arquivo: 

file_end = rf'C:\Users\{matricula}\Downloads\Tabela_data.csv'
file_aux_end = rf'C:\Users\{matricula}\Downloads\Tabela_data1.csv'

#%% Trabalhando o arquivo CSV: 


data1 = []
hora = []
submercado = []
preço = []

#Tira o cabeçalho:
    
with open(file_end) as f:
    with open(file_aux_end,'w') as f1:
        next(f, None) # skip header line
        for line in f:
            f1.write(line)

f = open( rf'C:\Users\{matricula}\Downloads\Tabela_data1.csv')

# Organiza os dados:
data = []
i=0
for line in f:
    data_line = line.rstrip().split(';')
    data.append(data_line)
    
    
    data1.append(data_line[0])
    hora.append(data_line[1])
    submercado.append(data_line[3])
    preço.append(data_line[4])
    tab_final =[data1,hora,submercado,preço]

f.close()

#%% Conectando com o Oracle:

try:
    con = cx_Oracle.connect(
        user="MR",
        password="my#MSvexa3",
        dsn="192.168.35.221:1541/corpp.cpfl.com.br")
    print("Connection sucessful")
except Exception as err:
    print("Error while creating the connection", err)
    
#Alterando o nome da base:


#%% Funções para carregar na base:

def rename_table(data_hoje,data_ontem):
    # construct an insert statement that add a new row to table
    sql = ("ALTER TABLE MRMI_PLD_HORA_{} RENAME TO MRMI_PLD_HORA_{}".format(data_ontem,data_hoje))
    print(sql)

    try:
        # establish a new connection
        with cx_Oracle.connect(
                user="MR",
                password="my#MSvexa3",
                dsn="192.168.35.221:1541/corpp.cpfl.com.br") as connection:
            # create a cursor
            with connection.cursor() as cursor:
                # execute the insert statement
                cursor.execute(sql)
                # commit work
                connection.commit()
    except cx_Oracle.Error as error:
        print('Error occurred:')
        print(error)
        
def insert_pld_att(data, hora, submercado, preço):
    # construct an insert statement that add a new row to table
    sql = ('insert into mrmi_pld_horario_att_diario(data, hora, submercado, preço) '
        'values(:data,:hora,:submercado,:preço)')

    try:
        # establish a new connection
        with cx_Oracle.connect(
                user="MR",
                password="my#MSvexa3",
                dsn="192.168.35.221:1541/corpp.cpfl.com.br") as connection:
            # create a cursor
            with connection.cursor() as cursor:
                # execute the insert statement
                cursor.execute(sql, [data, hora, submercado, preço])
                # commit work
                connection.commit()
    except cx_Oracle.Error as error:
        print('Error occurred:')
        print(error)

def insert_pld_dir(data, hora, submercado, preço):
    # construct an insert statement that add a new row to table
    sql = ("insert into mrmi_pld_hora_{}(data, hora, submercado, preço)"
"values(:data,:hora,:submercado,:preço)").format(data_hoje)

    try:
        # establish a new connection
        with cx_Oracle.connect(
                user="MR",
                password="my#MSvexa3",
                dsn="192.168.35.221:1541/corpp.cpfl.com.br") as connection:
            # create a cursor
            with connection.cursor() as cursor:
                # execute the insert statement
                cursor.execute(sql, [data, hora, submercado, preço])
                # commit work
                connection.commit()
    except cx_Oracle.Error as error:
        print('Error occurred:')
        print(error)

#%% Main:

if __name__ == '__main__':
    rename_table(data_hoje,data_ontem)

    for i in range(0,96):
        insert_pld_att(data1[i],hora[i], submercado[i], preço[i])
        insert_pld_dir(data1[i],hora[i], submercado[i], preço[i])
#%%

# Deletando os arquivos auxiliares:
file = rf'C:\Users\{matricula}\Downloads\Tabela_data.csv'
if(os.path.exists(file) and os.path.isfile(file)): 
  os.remove(file) 
  print("file Tabela_data deleted") 
else: 
  print("file Tabela_data not found") 
  
  
file = rf'C:\Users\{matricula}\Downloads\Tabela_data1.csv'
if(os.path.exists(file) and os.path.isfile(file)): 
  os.remove(file) 
  print("file Tabela_data1 deleted") 
else: 
  print("file Tabela_data1 not found") 
  
# Confirmação de execução:

data_atual = datetime.now()
print("Data update sucessfully on:",data_atual)

#%% Email Automático de Confirmação:
    
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'lucas.tavares@cpfl.com.br'
mail.Subject = 'Atualização Automática PLD Hora'
mail.Body = "Confirmação de Atualização da Base: {}".format(data_e_hora)

mail.Send()
