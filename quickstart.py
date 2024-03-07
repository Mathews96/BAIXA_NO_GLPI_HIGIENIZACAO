import os.path
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

#[ ATRIBUTOS ] <-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

# # If modifying these scopes, delete the file token.json.
#.readonly (para deixar somente leitura)
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1CJXTTrmaWSNTQnKD6oeLw9lGxEBNjz05S50L0MPLNLA"
SAMPLE_RANGE_NAME = "BD_HIGIENIZACAO!A2:D"

navegador = webdriver.Chrome()
campo_login_inicial = '//*[@id="login_name"]'
campo_senha_inicial = '/html/body/div[1]/div/div/div[1]/div/form/div/div[1]/div[4]/input'
botao_entrar_inicial = "/html/body/div[1]/div/div/div[1]/div/form/div/div[1]/div[7]/button"

#[ MÉTODOS ] <-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
def logarGLPI():
  #verificar se a página carregou por completo, se não, refresh()
  if navegador.find_elements(by=By.XPATH, value='//*[@id="login_name"]'):
    navegador.find_element(By.XPATH, campo_login_inicial).send_keys("12038463476")
    navegador.find_element(By.XPATH, campo_senha_inicial).send_keys('1796')
    navegador.find_element(By.XPATH, botao_entrar_inicial).click()

def tokenGoogleSheetsAPI():
  creds = None
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credencialgooglesheets.json", SCOPES
      )
      creds = flow.run_local_server(port=0)

    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())   
  return creds   


def getPlanilhaGeral():      
  try:    
    service = build("sheets", "v4", credentials=tokenGoogleSheetsAPI())

    #Ler informações [Células] o Google Sheets
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                                range=SAMPLE_RANGE_NAME).execute()     
    return result['values']

  except HttpError as err:
    return print(err)
  
  
  
def setCelulaPlanilha(aba, celula, valor): 
  try:
    service = build("sheets", "v4", credentials=tokenGoogleSheetsAPI())
    #Inserir / editar uma informação [Célula] no Google Sheets
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                                range=aba+celula, valueInputOption="USER_ENTERED", 
                                body={'values': valor}).execute()                

  except HttpError as err:
    print(err)  


#[ MAIN ] <-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

navegador.get("https://novoglpi.sms.maceio.al.gov.br/index.php?noAUTO=1")
time.sleep(10)

logarGLPI()

for i, chamadoID in enumerate(getPlanilhaGeral()):
  if(i > 0):
    linha = 0
    if(chamadoID[3] == "SIM"):
      print("Linha já conferida:"+ chamadoID[9], i)
      continue

    elif(chamadoID[3] == "NÃO"):
      print("Entrando no chamado do GLPI de ID: "+chamadoID[0], i)
      navegador.get("https://novoglpi.sms.maceio.al.gov.br/front/ticket.form.php?id="+chamadoID[0])
      if(navegador.find_elements(By.XPATH, "//*[contains(text(), 'Item não encontrado')]")):
        continue

      time.sleep(10)

      validandoCNS = False
      numeroCNS = 00000000000000

      tentativas = 0
      while(validandoCNS == False):
        numeroCNS = navegador.find_element(By.NAME, "cnfield").get_attribute('value')
        statusHigienizacao = navegador.find_element(By.NAME, "plugin_fields_statushigienizaofielddropdowns_id").get_attribute('value')
        print("STATUS HIGIENIZAÇÃO:"+ statusHigienizacao)
        #navegador.find_element(By.NAME, "plugin_fields_statushigienizaofielddropdowns_id").find_element(By.CLASS_NAME, "select2 select2-container select2-container--default").click()
        #navegador.find_element(By.NAME, "plugin_fields_statushigienizaofielddropdowns_id").click()
        statusHigienizacao = navegador.find_element(By.NAME, "plugin_fields_statushigienizaofielddropdowns_id")
        selecione = Select(statusHigienizacao)
        selecione.select_by_visible_text('Chamado higienizado - ')


        statusHigienizacao = navegador.find_element(By.NAME, "plugin_fields_statushigienizaofielddropdowns_id").get_attribute('value')
        print("STATUS HIGIENIZAÇÃO:"+ statusHigienizacao)

        print("STATUS HIGIENIZAÇÃO:"+ statusHigienizacao)
        if(str(numeroCNS) == str(chamadoID[2])):
          print("CNS correto: "+chamadoID[2], numeroCNS)
          validandoCNS = True
        elif(str(numeroCNS) != str(chamadoID[2])):
          print("CNS diferente: "+chamadoID[2], numeroCNS)          
          validandoCNS = False        
          navegador.refresh()
          tentativas+=1
          time.sleep(5)

          if(tentativas == 3):
            tentativas = 0 
            break

          

      if(navegador.find_elements(By.XPATH, value='//*[@id="page"]/div/div/div[2]/div[1]/h3')):
        valor = navegador.find_element(By.XPATH, '//*[@id="page"]/div/div/div[2]/div[1]/h3').text

        if(valor.find(chamadoID[2]) != -1):
          print('Entrou no chamado do GLPI id:'+chamadoID[0])
          time.sleep(10)

          if(navegador.find_elements(By.XPATH, '/html/body/div[14]')):
            print('Página carregada por completo!')
            dataAberturaChamado = navegador.find_element(By.XPATH, '//*[@id="plugin_fields_container_1853565001"]/div[10]/div/div/span[1]/span[1]/span').get_attribute('value')
            print(dataAberturaChamado)

            chamadoStatus = navegador.find_element(By.XPATH, '//*[@id="heading-main-item"]/button/span[1]/i').get_attribute('data-bs-original-title')
            print(chamadoStatus)

            dataRealizacao = navegador.find_element(By.NAME, "datarealizaofield").get_attribute('value')
            print(dataRealizacao)     

            marcado = "NÃO" if(not dataRealizacao.strip()) else "SIM"
            print(marcado)

            prestador = navegador.find_element(By.NAME, 'prestadorfield').get_attribute('value')
            print(prestador)

            linha = i+1
            #print("Dados inseridos na Aba: BD_HIGIENIZACAO! - Linha: D"+str(linha))
            setCelulaPlanilha('BD_HIGIENIZACAO!', 'D'+str(linha), [["SIM"]])
          else:
            print("Página não carregou por completo, atualizando-a")      
            #Refresh e verificar
            
        else:
          print('Não entrou no chamado do GLPI id:'+chamadoID[2])
          #Atualizar para o mesmo link e voltar ao início do código
        print(valor)
    else:
      print("Linha deu erro:"+ chamadoID[9], i)
      setCelulaPlanilha('BD_HIGIENIZACAO!', 'D'+str(linha), [["ERRO"]])
      continue

print('Programa Finalizado')