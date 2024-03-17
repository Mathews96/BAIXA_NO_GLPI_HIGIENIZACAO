import os.path
import time
import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC

#[ ATRIBUTOS ] <-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

# # If modifying these scopes, delete the file token.json.
#.readonly (para deixar somente leitura)
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1CJXTTrmaWSNTQnKD6oeLw9lGxEBNjz05S50L0MPLNLA"
SAMPLE_RANGE_NAME = "BD_HIGIENIZACAO!A2:G"

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
  linha = i+2
  if(linha > 0):    
    if(chamadoID[4] == "SIM"):
      print("Chamado já conferida:"+ chamadoID[1], linha)
      continue

    elif(chamadoID[4] == "NÃO"):
      print("Entrando no chamado do GLPI de ID: "+chamadoID[0], linha)
      navegador.get("https://novoglpi.sms.maceio.al.gov.br/front/ticket.form.php?id="+chamadoID[0])
      if(navegador.find_elements(By.XPATH, "//*[contains(text(), 'Item não encontrado')]")):
        continue

      time.sleep(10)

      validandoCNS = False
      numeroCNS = 00000000000000

      tentativas = 0
      while(validandoCNS == False):    
        WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.NAME, "cnfield")))

        try:
          numeroCNS = navegador.find_element(By.NAME, "cnfield").get_attribute('value').strip()
          print("-> DADOS DA PLANILHA - LINHA: "+str(linha)+"\nID_GLPI: "+chamadoID[0]+"\nNOME_PACIENTE: "+chamadoID[1]+"\nCNS: "+chamadoID[2]+"\nSTATUS HIGIENIZAÇÃO: "+chamadoID[3]+"\
          \nCHAMADO CONFERIDO: "+chamadoID[4]+"\nCHAMADO CONFERINDO EM: "+chamadoID[5]+"\nSTATUS DO CHAMADO: "+chamadoID[6])
          if(str(numeroCNS) == str(chamadoID[2]).strip()):
            print("CNS correto: "+chamadoID[2], numeroCNS)
            validandoCNS = True
          elif(str(numeroCNS) != str(chamadoID[2]).strip()):
            print("CNS diferente: "+chamadoID[2], numeroCNS)          
            validandoCNS = False        
            navegador.refresh()
            tentativas+=1
            time.sleep(5)

            if(tentativas == 3):
              tentativas = 0 
              break         
             
        except TimeoutException:
          continue          

      if(navegador.find_elements(By.XPATH, value='//*[@id="page"]/div/div/div[2]/div[1]/h3')):
        valor = navegador.find_element(By.XPATH, '//*[@id="page"]/div/div/div[2]/div[1]/h3').text

        if(valor.find(chamadoID[0]) != -1):
          print('Entrou no chamado do GLPI id:'+chamadoID[0])
          
          dataHoraAtual = datetime.datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
          print(str(dataHoraAtual))
      
          try:          
            WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, "//*[contains(text(), 'Data de Nascimento')]")))
            '''chamadoStatus = "Nenhum"
            if navegador.find_element(By.XPATH, "//*[@class='me-1']//*[@data-bs-original-title='Novo']'"):
              chamadoStatus = "Novo"

            elif navegador.find_element(By.XPATH, "//*[@class='me-1']//*[@data-bs-original-title='Em atendimento (atribuído)']"):
              chamadoStatus = "Em atendimento (atribuído)"

            elif navegador.find_element(By.XPATH, "//*[@class='me-1']//*[@data-bs-original-title='Em atendimento (Planejado)']"):
              chamadoStatus = "Em atendimento (Planejado)"

            elif navegador.find_element(By.XPATH, "//*[@class='me-1']//*[@data-bs-original-title='Pendente']"):
              chamadoStatus = "Pendente"

            elif navegador.find_element(By.XPATH, "//*[@class='me-1']//*[@data-bs-original-title='Solucionado']"):
              chamadoStatus = "Solucionado"

            elif navegador.find_element(By.XPATH, "//*[@class='me-1']//*[@data-bs-original-title='Fechado']"):
              chamadoStatus = "Fechado"
            else:
              chamadoStats = "ERRO-STATUS"'''
              
            if navegador.find_elements(By.XPATH, "//*[contains(@title, 'Novo')]") or navegador.find_elements(By.XPATH, "//*[contains(@title, 'Em atendimento (atribuído)')]")\
            or navegador.find_elements(By.XPATH, "//*[contains(@title, 'Em atendimento (Planejado)')]") or navegador.find_elements(By.XPATH, "//*[contains(@title, 'Pendente')]"):
          
              #navegador.execute_script("document.body.style.zoom = '50%';")
              #navegador.execute_script("arguments[0].scrollIntoView();", navegador.find_element(By.XPATH, "//*[@name='plugin_fields_gestantefielddropdowns_id']"))
              #time.sleep(10)

              try:
                campoStatusHigienizao = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, "//*[@data-select2-id='47']")))
                campoStatusHigienizao.click()
              except Exception as e:
                print("[ERRO]AO ABRIR A LISTA DOS STATUS DE HIGIENIZACAO: ", e)


              try:
                selecionandoItem = WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, "//*[@title='"+chamadoID[3]+"']")))
                selecionandoItem.click()

              except Exception as e:
                print("[ERRO]AO CLICAR NO STATUS DA LISTA HIGIENIZACAO: ", e)
                continue

              campoWhatsApp = navegador.find_element(By.NAME, "whatsappfield")
              if campoWhatsApp.get_attribute('value') == "": 
                campoWhatsApp.send_keys("00000000000")

              navegador.find_element(By.NAME, "update").click()    

              setCelulaPlanilha('BD_HIGIENIZACAO!', 'E'+str(linha), [["SIM"]])
              setCelulaPlanilha('BD_HIGIENIZACAO!', 'F'+str(linha), [[str(dataHoraAtual)]])
              print("[STATUS DO CHAMADO: ABERTO]LINHA ALTERADA: ", linha)

            else:
              setCelulaPlanilha('BD_HIGIENIZACAO!', 'E'+str(linha), [["SIM"]])
              setCelulaPlanilha('BD_HIGIENIZACAO!', 'F'+str(linha), [[str(dataHoraAtual)]])
              print("[STATUS DO CHAMADO: FECHADO / SOLUCIONADO]LINHA ALTERADA: ", linha)

          except TimeoutException:          
            print('Não entrou no chamado do GLPI id:'+chamadoID[2])
            setCelulaPlanilha('BD_HIGIENIZACAO!', 'E'+str(linha), [["ERRO AO CARREGAR A PÁGINA"]]) 
            setCelulaPlanilha('BD_HIGIENIZACAO!', 'F'+str(linha), [[str(dataHoraAtual)]])
            continue

print('Chamados atualizados com sucesso! Por: mathewsfreire@gmail.com')