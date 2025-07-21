from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import ElementNotInteractableException, TimeoutException, NoSuchElementException

import pandas as pd
import time


def ir_para_url(driver, url):
    """
    Direciona o navegador para a URL especificada.
    """
    driver.get(url)

def preencher_data(driver, wait, id_campo, data):
    """
    Preenche um campo de data com formatação, garantindo que o cursor comece do início.
    """
    campo = wait.until(EC.presence_of_element_located((By.ID, id_campo)))
    campo.click()
    campo.send_keys(Keys.HOME, data)

def buscar_nota(driver, wait, cpf, valor_esperado):
    """
    Realiza a busca da nota por CPF e verifica o valor.
    Retorna 'sim' se valor encontrado for igual, 'não' caso contrário.
    """
    # Aguarda que não haja overlay de popup visível
    try:
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "ui-widget-overlay")))
    except:
        pass

        # Preenche o campo CPF
    campo_cpf = wait.until(EC.presence_of_element_located((By.ID, "MainContentPlaceHolder_txtCnpjCpfTomador")))
    campo_cpf.clear()
    campo_cpf.send_keys(cpf)

    # Clica no botão Pesquisar
    botao_pesquisar = wait.until(EC.element_to_be_clickable((By.ID, "btnPesquisar")))
    botao_pesquisar.click()

    # Espera um pouco para o sistema responder
    time.sleep(1.0)

    # Caso tenha mensagem de erro, trata e retorna "não"
    if fechar_popup_se_existir(driver, wait):
        return "não"

    # Verifica se apareceu múltiplos registros e tenta clicar, caso esteja visível e clicável
    try:
        btn_endereco = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.ID,
                                        "MainContentPlaceHolder_WCListaRegistro_rptPesquisaDocumento_EditarGuiaImageButton_0"))
        )
        btn_endereco.click()
        time.sleep(1.5)

        # Após selecionar o endereço, pode surgir novamente o popup de "não encontrado"
        if fechar_popup_se_existir(driver, wait):
            return "não"

    except (NoSuchElementException, TimeoutException, ElementNotInteractableException):
        pass

    # Tenta capturar todas as notas exibidas
    try:
        elementos_valor = driver.find_elements(By.XPATH,
                                               "//a[starts-with(@id, 'MainContentPlaceHolder_rptDados_LinkButton7_')]")
        for el in elementos_valor:
            texto = el.text.strip().replace("R$", "").replace(",", ".")
            try:
                valor_site = float(texto)
                if abs(valor_site - float(valor_esperado)) < 0.01:
                    return "sim"
            except ValueError:
                continue  # caso algum valor seja inválido
        return "não"  # Nenhuma nota bateu
    except:
        return "não"

def fechar_popup_se_existir(driver, wait):
    """
    Tenta fechar o pop-up de mensagem, se estiver visível.
    """
    try:
        # Espera até 2 segundos para ver se o botão do popup aparece
        botao_fechar = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.ID, "btn-mensagem-dialog"))
        )
        botao_fechar.click()
        # Espera o overlay sumir
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "ui-widget-overlay")))
        # Espera um pouco para o sistema responder
        time.sleep(1.0)
        return True
    except TimeoutException:
        return False

# Data das notas a serem procuradas
data = "18/07/2025"

# Caminho para seu chromedriver
caminho_driver = r"C:\Users\pc\Desktop\chromedriver\chromedriver-win64\chromedriver.exe"

# Lê a planilha garantindo CPF como texto e deixa com 11 digitos
df = pd.read_excel(r"C:\Users\pc\Desktop\planilha notas\nfe.xlsx")
df["cpf"] = df["cpf"].fillna("").astype(str).str.zfill(11)

# Inicia o navegador
servico = Service(caminho_driver)
driver = webdriver.Chrome(service=servico)
driver.maximize_window()
wait = WebDriverWait(driver, 15)

# 1) Abre o site de login
ir_para_url(driver, "https://nfe.osasco.sp.gov.br/EissnfeWebApp/Portal/Default.aspx")

# 2) Faz login
usuario_input = wait.until(EC.presence_of_element_located((By.ID, "UsuarioTextBox")))
senha_input   = wait.until(EC.presence_of_element_located((By.ID, "SenhaTextBox")))
usuario_input.send_keys("-")
senha_input.send_keys("-")

botao_login = wait.until(EC.element_to_be_clickable((By.ID, "btnEntrar")))
botao_login.click()

# 3) Pequena espera para carregar
time.sleep(1.5)

# 4) Vai direto para a página de pesquisa
ir_para_url(driver, "https://nfe.osasco.sp.gov.br/EissnfeWebApp/Sistema/Prestador/PesquisarNFE.aspx?IdPermissaoAcesso=132")

# 5) Coloca a data da pesquisa
preencher_data(driver, wait, "txtDataGeracaoInicio", f"{data}")
preencher_data(driver, wait, "txtDataGeracaoFim", f"{data}")

# 6) Loop
resultados = []

for index, row in df.iterrows():
    pagador = row["pagador"]
    cpf = row["cpf"]
    valor_nota = row["valor"]

    resultado = buscar_nota(driver, wait, cpf, valor_nota)
    print(f"{pagador} | CPF: {cpf} | Valor: {valor_nota} → {resultado}")
    resultados.append(resultado)

# Adiciona coluna 'resultado' ao DataFrame
df["resultado"] = resultados

# Salva o resultado final
df.to_excel(r"C:\Users\pc\Desktop\planilha notas\nfe_resultado.xlsx", index=False)
