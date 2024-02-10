from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
from credenciais import segredos

codigo = segredos.get('codigo')
senha = segredos.get('senha')
caminhoBD = r"C:\Users\naelm\Downloads\Telegram Desktop\bd2.xlsx"
primeiraLinha = 27

# Abrir o Chrome e fazer login
navegador = webdriver.Chrome()
navegador.get('https://sigo.sh.srv.br/pls/webmin/webnewcadastrousuario.login')
navegador.find_element("xpath", '//*[@id="webnewcadastrousuario"]/table/tbody/tr[1]/td[2]/table/tbody/tr[1]/td[2]/input').send_keys(codigo)
navegador.find_element("xpath", '//*[@id="webnewcadastrousuario"]/table/tbody/tr[1]/td[2]/table/tbody/tr[3]/td[2]/input').click()
time.sleep(2)
WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="pSenha"]')))
navegador.find_element("xpath",'//*[@id="pSenha"]').send_keys(senha)
navegador.find_element("xpath", '//*[@id="Prosseguir"]').click()
time.sleep(1)

# Ir para Inclusão de Titular
WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/table[2]/tbody/tr/td/p[5]/input')))
navegador.find_element("xpath", '/html/body/div[2]/div/div/table[2]/tbody/tr/td/p[5]/input').click()
WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/form/table/tbody/tr[5]/td/input[2]')))
navegador.find_element("xpath", '/html/body/div[2]/div/div/form/table/tbody/tr[5]/td/input[2]').click()
WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/table/tbody/tr/td[2]/table/tbody/tr[3]/td/div/ul/li[1]/a/i')))
navegador.find_element("xpath", '/html/body/div[2]/div/div/table/tbody/tr/td[2]/table/tbody/tr[3]/td/div/ul/li[1]/a/i').click()

# Coleta as informacoes da planilha
planilha = load_workbook(caminhoBD)
aba_ativa = planilha.active
for linha in aba_ativa.iter_rows(min_row=primeiraLinha, max_row=aba_ativa.max_row):
    celula = linha[0]
    if celula.value is not None:
        nome = aba_ativa[f"D{celula.row}"].value
        dtNascimento = aba_ativa[f"E{celula.row}"].value
        sexo = aba_ativa[f"F{celula.row}"].value
        cpfTitular = aba_ativa[f"G{celula.row}"].value
        cpfReal = aba_ativa[f"H{celula.row}"].value
        dependencia = aba_ativa[f"I{celula.row}"].value
        cep = aba_ativa[f"Y{celula.row}"].value
        print(cpfReal)
        # Acessar CPF
        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/table[3]/tbody/tr/td/p[2]/input')))
        navegador.find_element("xpath", '/html/body/div[2]/div/div/table[3]/tbody/tr/td/p[2]/input').clear()
        navegador.find_element("xpath", '/html/body/div[2]/div/div/table[3]/tbody/tr/td/p[2]/input').send_keys(cpfReal)
        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="id"]')))
        time.sleep(2)
        navegador.find_element("xpath", '//*[@id="id"]').click()
        navegador.find_element("xpath", '// *[ @ id = "pNomeTitular"]').send_keys(nome)
        #navegador.find_element("xpath", '// *[ @ id = "pCep"]').click()
        #navegador.find_element("xpath", '// *[ @ id = "pCep"]').send_keys(cep)
        #navegador.find_element("xpath", '// *[ @ id = "pDs_Ponto_Referencia"]').click()

        #Chama a pagina para o proximo
        time.sleep(10)
        navegador.back()
        # Grava a confirmacao na planilha
        #aba_ativa[f"K{celula.row}"] = "Jeff"
        #planilha.save(caminhoBD)

time.sleep(120)