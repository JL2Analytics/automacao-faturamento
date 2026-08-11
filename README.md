import os
import time
import logging
import openpyxl
from O365 import Account
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ==================================================
# CONFIGURAÇÃO DE LOGS (Substitui os Pop-ups de interface)
# ==================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# ==================================================
# CONFIGURAÇÕES DE AMBIENTE / VARIÁVEIS
# ==================================================
DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR", os.path.abspath("./downloads"))
ORIGEM = os.path.join(DOWNLOAD_DIR, "Veiculos Logisticos.xlsx")
DESTINO_LOCAL = os.path.abspath("./temp/Programacao_de_veiculos.xlsx")

# Credenciais do Portal Raster
USUARIO = os.getenv("RASTER_USER", "13216634646")
SENHA = os.getenv("RASTER_PASS", "123456")

# Garantir criação das pastas temporárias locais
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(os.path.dirname(DESTINO_LOCAL), exist_ok=True)

# ==================================================
# FUNÇÃO: UPLOAD PARA ONEDRIVE / SHAREPOINT (GRAPH API)
# ==================================================
def enviar_para_sharepoint(caminho_arquivo_local, pasta_destino):
    """
    Autentica na Microsoft Graph API usando os Secrets do Azure 
    e envia o arquivo para a pasta BI no OneDrive/SharePoint.
    """
    CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
    CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
    TENANT_ID = os.getenv("AZURE_TENANT_ID")

    # Se não houver credenciais do Azure configuradas, pula o envio em nuvem
    if not all([CLIENT_ID, CLIENT_SECRET, TENANT_ID]):
        logging.warning("Credenciais do Azure não encontradas. O upload para o SharePoint/OneDrive foi ignorado.")
        return

    credentials = (CLIENT_ID, CLIENT_SECRET)
    account = Account(credentials, auth_flow_type='credentials', tenant_id=TENANT_ID)

    if not account.authenticate():
        raise Exception("Falha na autenticação com a Microsoft Graph API.")

    logging.info("Autenticado no Microsoft 365 com sucesso.")

    # Conecta no OneDrive do usuário / Espaço corporativo
    storage = account.storage()
    drive = storage.get_default_drive()

    # Obtém a pasta de destino (Documents/BI)
    try:
        folder = drive.get_item_by_path(pasta_destino)
    except Exception:
        # Tenta o nome em português caso Documents falhe
        folder = drive.get_item_by_path("Documentos/BI")

    logging.info(f"Enviando {caminho_arquivo_local} para a pasta {pasta_destino} no OneDrive...")
    folder.upload_file(item=caminho_arquivo_local)
    logging.info("Upload para o OneDrive/SharePoint concluído com sucesso!")

# ==================================================
# FUNÇÕES MANIPULAÇÃO DE PLANILHA
# ==================================================
def copiar_dados_planilha(origem_path, destino_path):
    wb_origem = openpyxl.load_workbook(origem_path)
    ws_origem = wb_origem.active

    if ws_origem.max_row < 2:
        wb_origem.close()
        raise Exception("Planilha exportada está vazia ou incompleta.")

    dados = [list(row) for row in ws_origem.iter_rows(values_only=True)]

    if os.path.exists(destino_path):
        wb_destino = openpyxl.load_workbook(destino_path)
        ws_destino = wb_destino.active
        ws_destino.delete_rows(1, ws_destino.max_row)
    else:
        wb_destino = openpyxl.Workbook()
        ws_destino = wb_destino.active

    for i, linha in enumerate(dados, start=1):
        for j, valor in enumerate(linha, start=1):
            ws_destino.cell(row=i, column=j, value=valor)

    wb_destino.save(destino_path)
    wb_origem.close()
    wb_destino.close()

    # Validação final
    wb_check = openpyxl.load_workbook(destino_path)
    ws_check = wb_check.active
    if ws_check.max_row < 2:
        wb_check.close()
        raise Exception("Falha ao gravar dados no arquivo destino.")
    wb_check.close()

    os.remove(origem_path)

def esperar_download(caminho, timeout=180):
    for _ in range(timeout):
        if os.path.exists(caminho) and not caminho.endswith(".crdownload"):
            return True
        time.sleep(1)
    return False

# ==================================================
# FUNÇÕES SELENIUM
# ==================================================
def fechar_alerta_se_existir(driver, wait):
    try:
        alerta = wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'dx-popup-content')]"))
        )
        for texto in ["OK", "Fechar", "Entendi"]:
            try:
                alerta.find_element(By.XPATH, f".//span[text()='{texto}']").click()
                time.sleep(2)
                return
            except:
                pass
    except:
        pass

def aguardar_overlay_sumir(wait):
    try:
        wait.until(
            EC.invisibility_of_element_located((By.XPATH, "//div[contains(@class,'dx-overlay-shader')]"))
        )
    except:
        pass

def clicar_por_texto(wait, driver, texto):
    aguardar_overlay_sumir(wait)
    el = wait.until(
        EC.element_to_be_clickable((By.XPATH, f"//span[normalize-space(text())='{texto}']"))
    )
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(1)
    driver.execute_script("arguments[0].click();", el)

def exportar_todos_os_dados(wait, driver):
    aguardar_overlay_sumir(wait)
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "dx-datagrid-rowsview")))
    time.sleep(2)

    container = wait.until(
        EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'dx-datagrid-export-button')]"))
    )
    spindown = container.find_element(By.XPATH, ".//i[contains(@class,'dx-icon-spindown')]")
    driver.execute_script("arguments[0].click();", spindown)
    time.sleep(2)

    opcao = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//div[contains(@class,'dx-list-item-content') and contains(normalize-space(.), 'Exportar todos os dados')]")
        )
    )
    driver.execute_script("arguments[0].click();", opcao)

# ==================================================
# EXECUÇÃO PRINCIPAL
# ==================================================
if __name__ == "__main__":
    logging.info("Atualização iniciada. Processando o Status dos veículos...")

    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-extensions")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
    )

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    # Permissão para downloads em modo Headless
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": DOWNLOAD_DIR}
    )

    wait = WebDriverWait(driver, 40)

    try:
        driver.get(
            "https://auth.rastergr.com.br/realms/raster/protocol/openid-connect/auth"
            "?client_id=log&response_type=code&redirect_uri=https://www.rasterlog.com.br"
        )

        wait.until(EC.presence_of_element_located((By.ID, "username")))
        driver.find_element(By.ID, "username").send_keys(USUARIO)
        driver.find_element(By.ID, "password").send_keys(SENHA + Keys.ENTER)

        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "dx-popup-content")))

        tagbox = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'dx-tagbox')]")))
        driver.execute_script("arguments[0].click();", tagbox)
        time.sleep(2)

        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'dx-list-select-all-checkbox')]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Prosseguir']"))).click()

        time.sleep(3)
        fechar_alerta_se_existir(driver, wait)

        clicar_por_texto(wait, driver, "Visões")
        clicar_por_texto(wait, driver, "Veículos logisticos")

        time.sleep(5)
        exportar_todos_os_dados(wait, driver)

        if not esperar_download(ORIGEM):
            raise Exception("Arquivo de exportação não foi encontrado ou não concluiu o download.")

        # 1. Trata a planilha localmente
        copiar_dados_planilha(ORIGEM, DESTINO_LOCAL)
        logging.info(f"Dados salvos temporariamente em: {DESTINO_LOCAL}")

        # 2. Envia para a pasta BI no OneDrive / SharePoint
        enviar_para_sharepoint(
            caminho_arquivo_local=DESTINO_LOCAL,
            pasta_destino="Documents/BI"
        )

        logging.info("Processo concluído com sucesso!")

    except Exception as e:
        logging.error(f"Ocorreu um erro durante a execução: {e}", exc_info=True)

    finally:
        try:
            driver.quit()
        except:
            pass
