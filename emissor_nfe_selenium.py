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

# Função para fechar o pop-up se aparecer
def fechar_popup_se_existir(driver):
    try:
        botao_fechar = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='Fechar']/.."))
        )
        botao_fechar.click()
    except TimeoutException:
        pass  # Nenhum pop-up apareceu

# Caminho para seu chromedriver
caminho_driver = r"C:\Users\pc\Desktop\chromedriver\chromedriver-win64\chromedriver.exe"

# Lê a planilha garantindo CPF como texto e deixa com 11 digitos
df = pd.read_excel(r"C:\Users\pc\Desktop\planilha notas\nfe.xlsx")
df["cpf"] = df["cpf"].fillna("").astype(str).str.zfill(11)

# Passo a Passo
# 1.0 - Iniciar o navegador
servico = Service(caminho_driver)
driver = webdriver.Chrome(service=servico)
wait = WebDriverWait(driver, 15)

# 2.0 - Abrir o site de login
driver.get("https://nfe.osasco.sp.gov.br/EissnfeWebApp/Portal/Default.aspx")

# 3.0 - Fazer login
usuario_input = wait.until(EC.presence_of_element_located((By.ID, "UsuarioTextBox")))
senha_input   = wait.until(EC.presence_of_element_located((By.ID, "SenhaTextBox")))
usuario_input.send_keys("NFE-")
senha_input.send_keys("PCE-")
botao_login = wait.until(EC.element_to_be_clickable((By.ID, "btnEntrar")))
botao_login.click()

# 3.1 - Esperar o site carregar
time.sleep(1.5)

# 4.0 - Ir para a página de emissão
driver.get("https://nfe.osasco.sp.gov.br/EissnfeWebApp/Sistema/Prestador/EmitirNFE.aspx?IdPermissaoAcesso=131")

# --- LOOP PELO EXCEL ---
for index, row in df.iterrows():
    pagador = row["pagador"]
    cpf = row["cpf"]
    valor_nota = row["valor"]

    # 5.0 - Inserir CPF
    documento_input = wait.until(
        EC.presence_of_element_located((By.ID, "DocumentoTextBox"))
    )
    documento_input.clear()
    documento_input.send_keys(cpf)

    # 5.1 - Pesquisar CPF
    botao_pesquisar = wait.until(EC.element_to_be_clickable((By.ID, "PesquisarTomadorButton")))
    botao_pesquisar.click()

    # 6.0 - Verificação de Pop-up
    try:
        # 6.1 - Pop-Up de seleção de endereço
        botao_selecionar_endereco = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable(
                (By.ID, "MainContentPlaceHolder_MainContentPlaceHolder_ContribuintesRepeater_EditarImageButton_0"))
        )
        botao_selecionar_endereco.click()

        # 6.2 - Pop-up extra pós seleção de endereço
        fechar_popup_se_existir(driver)

    except TimeoutException:
        try:
            # 6.3 - Pop-up de falta de cadastro
            botao_fechar = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='Fechar']/.."))
            )
            botao_fechar.click()

            # 6.4 - Preenchimento do nome do cadastro
            nome_input = wait.until(
                EC.presence_of_element_located(
                    (By.ID, "MainContentPlaceHolder_MainContentPlaceHolder_NomeTomadorTextBox"))
            )
            nome_input.clear()
            nome_input.send_keys(pagador)

        except TimeoutException:
            # Nenhum dos dois pop-ups apareceu — segue normalmente
            pass


    # 7.0 - Preenchimento do campo da aliquota
    aliquota_input = wait.until(EC.element_to_be_clickable((By.ID, "AliquotaISSTextBox")))
    # 7.1 - Selecionar tudo e apaga
    aliquota_input.send_keys(Keys.CONTROL, "a")
    aliquota_input.send_keys(Keys.BACKSPACE)
    # 7.2 - Digitar o novo valor e confirma
    aliquota_input.send_keys("2,00")
    aliquota_input.send_keys(Keys.ENTER)

    # 8.0 - Preenchimento do campo descrição
    try:
        descricao_input = wait.until(
            EC.presence_of_element_located((By.ID,
                                            "MainContentPlaceHolder_MainContentPlaceHolder_EmissaoNFEWebUserControl1_OutrasInformacoesTextBox"))
        )
        descricao_input.clear()
        descricao_input.send_keys(
            "Serviços de Impressões, cópias e congêneres em geral. "
            "\nRealizado em 12/04/2025 "
            "\nNota fiscal emitida pelo sistema."
        )
    except StaleElementReferenceException:
        # 8.1 - Em caso de erro relocaliza caso o elemento tenha se tornado obsoleto
        descricao_input = driver.find_element(By.ID,
                                              "MainContentPlaceHolder_MainContentPlaceHolder_EmissaoNFEWebUserControl1_OutrasInformacoesTextBox")
        descricao_input.clear()
        descricao_input.send_keys(
            "Serviços de Impressões, cópias e congêneres em geral. "
            "\nRealizado em 12/04/2025 "
            "\nNota fiscal emitida pelo sistema."
        )

    # 9.0 - Desmarcar a checkbox de e-mail, se estiver ativada
    try:
        checkbox_email = wait.until(
            EC.presence_of_element_located((By.ID, "NotificarTomadorPorEmailCheckBox"))
        )
        if checkbox_email.is_selected():
            checkbox_email.click()
    except Exception as e:
        print(f"Erro ao tentar desmarcar checkbox de e-mail: {e}")
    time.sleep(1.0)

    # 10.0 - Preenchimento do valor da nota
    valor_str = f"{valor_nota:.2f}".replace(".", ",")

    # 10.1 - campo do ValorTotalNotaTextBox
    valor_input = wait.until(EC.presence_of_element_located((By.ID, "ValorTotalNotaTextBox")))
    valor_input.send_keys(Keys.CONTROL, "a")
    valor_input.send_keys(Keys.BACKSPACE)
    valor_input.send_keys(valor_str)
    valor_input.send_keys(Keys.ENTER)
    time.sleep(1.5)

    # 11.0 Emissão da nota
    emitir_btn = wait.until(EC.element_to_be_clickable((By.ID, "ConfirmarVisualizarButton")))
    emitir_btn.click()
    time.sleep(1.5)

    # 12.0 - Fechar Pop-up
    abas = driver.window_handles
    if len(abas) > 1:
        driver.switch_to.window(abas[-1])
        driver.close()
        driver.switch_to.window(abas[0])


# Fim do Loop
print('Dia 12 Finalizado')
driver.quit()