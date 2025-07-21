from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import StaleElementReferenceException
import pandas as pd
import time

# 1) Função para fechar o pop-up se aparecer
def fechar_popup_se_existir(driver):
    try:
        botao_fechar = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='Fechar']/.."))
        )
        botao_fechar.click()
    except TimeoutException:
        pass  # Nenhum pop-up apareceu

# 2) Caminho para seu chromedriver
caminho_driver = r"C:\Users\pc\Desktop\chromedriver\chromedriver-win64\chromedriver.exe"

# 3) Lê a planilha garantindo CPF como texto e deixa com 11 dígitos
df = pd.read_excel(r"C:\Users\pc\Desktop\planilha notas\nfe.xlsx")
df["cpf"] = df["cpf"].fillna("").astype(str).str.zfill(11)

# 4) Inicia o navegador
servico = Service(caminho_driver)
driver = webdriver.Chrome(service=servico)
driver.maximize_window()
wait = WebDriverWait(driver, 15)

# 5) Abre o site de login
driver.get("https://nfe.osasco.sp.gov.br/EissnfeWebApp/Portal/Default.aspx")

# 6) Faz login
usuario_input = wait.until(EC.presence_of_element_located((By.ID, "UsuarioTextBox")))
senha_input   = wait.until(EC.presence_of_element_located((By.ID, "SenhaTextBox")))
usuario_input.send_keys("-")
senha_input.send_keys("-")

botao_login = wait.until(EC.element_to_be_clickable((By.ID, "btnEntrar")))
botao_login.click()

# 7) Pequena espera para carregar
time.sleep(1.5)

# 8) Vai direto para a página de emissão
driver.get("https://nfe.osasco.sp.gov.br/EissnfeWebApp/Sistema/Prestador/EmitirNFE.aspx?IdPermissaoAcesso=131")

# 9) Loop pelos dados do Excel
for index, row in df.iterrows():
    pagador = row["pagador"]
    cpf = row["cpf"]
    valor_nota = row["valor"]

    # 9.1) Preenche o CPF do tomador
    documento_input = wait.until(EC.presence_of_element_located((By.ID, "DocumentoTextBox")))
    documento_input.clear()
    documento_input.send_keys(cpf)

    # 9.2) Clica em Pesquisar
    botao_pesquisar = wait.until(EC.element_to_be_clickable((By.ID, "PesquisarTomadorButton")))
    botao_pesquisar.click()

    # 9.3) Trata seleção de endereço ou preenchimento manual
    try:
        botao_selecionar_endereco = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable(
                (By.ID, "MainContentPlaceHolder_MainContentPlaceHolder_ContribuintesRepeater_EditarImageButton_0"))
        )
        botao_selecionar_endereco.click()
        fechar_popup_se_existir(driver)

    except TimeoutException:
        try:
            botao_fechar = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='Fechar']/.."))
            )
            botao_fechar.click()

            nome_input = wait.until(EC.presence_of_element_located(
                (By.ID, "MainContentPlaceHolder_MainContentPlaceHolder_NomeTomadorTextBox")))
            nome_input.clear()
            nome_input.send_keys(pagador)
        except TimeoutException:
            pass

    # 9.4) Preenche alíquota ISS
    aliquota_input = wait.until(EC.element_to_be_clickable((By.ID, "AliquotaISSTextBox")))
    aliquota_input.send_keys(Keys.CONTROL, "a")
    aliquota_input.send_keys(Keys.BACKSPACE)
    aliquota_input.send_keys("2,00")
    aliquota_input.send_keys(Keys.ENTER)

    # 9.5) Preenche a descrição dos serviços
    try:
        descricao_input = wait.until(
            EC.presence_of_element_located((By.ID, "MainContentPlaceHolder_MainContentPlaceHolder_EmissaoNFEWebUserControl1_OutrasInformacoesTextBox"))
        )
        descricao_input.clear()
        descricao_input.send_keys(
            "Serviços de Impressões, cópias e congêneres em geral. "
            "\nRealizado em 16/07/2025"
            "\nNota fiscal emitida pelo sistema."
        )
    except StaleElementReferenceException:
        descricao_input = driver.find_element(By.ID, "MainContentPlaceHolder_MainContentPlaceHolder_EmissaoNFEWebUserControl1_OutrasInformacoesTextBox")
        descricao_input.clear()
        descricao_input.send_keys(
            "Serviços de Impressões, cópias e congêneres em geral. "
            "\nRealizado em 16/07/2025"
            "\nNota fiscal emitida pelo sistema."
        )

    # 9.6) Desmarca checkbox de e-mail
    try:
        checkbox_email = wait.until(EC.presence_of_element_located((By.ID, "NotificarTomadorPorEmailCheckBox")))
        if checkbox_email.is_selected():
            checkbox_email.click()
    except Exception as e:
        print(f"Erro ao tentar desmarcar checkbox de e-mail: {e}")
    time.sleep(1.5)

    # 9.7) Preenche o valor da nota
    valor_str = f"{valor_nota:.2f}".replace(".", ",")
    valor_input = wait.until(EC.presence_of_element_located((By.ID, "ValorTotalNotaTextBox")))
    valor_input.send_keys(Keys.CONTROL, "a")
    valor_input.send_keys(Keys.BACKSPACE)
    valor_input.send_keys(valor_str)
    valor_input.send_keys(Keys.ENTER)
    time.sleep(1.5)

    # 9.8) Clica no botão Emitir
    emitir_btn = wait.until(EC.element_to_be_clickable((By.ID, "ConfirmarVisualizarButton")))
    emitir_btn.click()
    time.sleep(1.5)

    # 9.9) Fecha aba extra, se aberta
    abas = driver.window_handles
    if len(abas) > 1:
        driver.switch_to.window(abas[-1])
        driver.close()
        driver.switch_to.window(abas[0])

# 10) Encerra execução
print('Dia 16 Finalizado')
driver.quit()
