import os

from selenium import webdriver

from selenium.common import exceptions
from selenium.webdriver.common import desired_capabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from pathlib import Path

default_path = os.getcwd()


def start(path=default_path, download_pdf=True, prompt_download=False, accept_insecure=True):
    """Abre o driver do Selenium com as configurações informadas.

    Args:
        path (str, optional): Sets Chrome default download path. Defaults to default_path.
        download_pdf (bool, optional): Download PDF instead of openning it
        internally.
        prompt_download (bool, optional): Prompt for download. Defaults to False.
        accept_insecure (bool, optional): Allow insecure connections. Defaults to True.

    Returns: WebDriver Class: classe do navegador Selenium
    """

    path = str(Path(path))

    desired_caps = desired_capabilities.DesiredCapabilities.CHROME.copy()
    desired_caps['acceptInsecureCerts'] = accept_insecure
    desired_caps['purge-memory-button'] = True

    options = webdriver.ChromeOptions()
    preferences = {'download.default_directory': path,  # caminho padrão
                   "download.prompt_for_download": prompt_download,  # baixa sem perguntar
                   'safebrowsing.enabled': False,  # baixa também executável
                   'plugins.always_open_pdf_externally': download_pdf}  # ,      #  salva PDF em vez de abrir
    # 'useAutomationExtension': False}
    
    options.add_experimental_option("prefs", preferences)
    options.add_experimental_option("excludeSwitches", ['enable-automation']) # desabilita o aviso de navegador automático
    options.add_argument('user-data-dir=C:\\temp\chrome_profile')
    options.add_argument('profile-directory="Profile 1"') #  perfil "AFR"
    # options.add_argument('--headless') 

    driver = webdriver.Chrome(
        options=options, desired_capabilities=desired_caps)
    driver.set_window_position(750, 0)
    driver.set_window_size(803, 831)

    return driver


'''def xpath(xpath):
    """xpath Atalho para o comando 'driver.find_element_by_xpath' do Selenium.
    
    Args:
        xpath ([str]): X-Path absoluto ou relativo do elemento
    
    
    Returns:
        [Selenium WebElement]: Classe de WebElement do Selenium.
    
    """
    return driver.find_element_by_xpath(xpath)
'''


def driver_ativo(driver):
    '''Verifica se a sessão do driver está ativa.
    Retorna True ou False'''

    # Não consegui achar um listener pra caso a janela do driver seja fechada.
    # Para verificar, criei uma tentativa de acesso arbitrária.
    # Se o driver estiver fechado, pega a exceção e retorna False.

    try:
        driver.title
        return True
    except:
        return False


if __name__ == '__main__':
    driver = start()
