import sys

from getpass import getpass
from pathlib import Path
from pprint import pprint as pprint
from webbrowser import Elinks

from selenium import webdriver

from selenium.common import exceptions
from selenium.webdriver.common import desired_capabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

import tkinter as tk

import browser

erro_de_quando_driver_fecha_aparentemente_sozinho = '''MaxRetryError(_pool, url, error or ResponseError(cause))
urllib3.exceptions.MaxRetryError'''

driver = browser.start()
# driver=webdriver.Chrome()
driver.minimize_window()


def q():
    driver.quit()
    sys.exit()


def login():

    def open_vpn():
        '''Abre acesso ao VPN em sessão comum do navegador (sem ser pelo
        Selenium), para evitar que a sessão seja encerrada caso o webdriver
        o seja'''

        import webbrowser
        vpn_url = 'https://vpnssl.fazenda.sp.gov.br/sslvpn/Portal/Main'
        chrome_path = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"

        if Path(chrome_path).exists():
            webbrowser.register(
                'chrome', None, webbrowser.BackgroundBrowser(chrome_path), preferred=True)
        else:
            print('''Erro. Navegador Chrome não encontrado.
            Acessando site do VPN pelo navegador padrão.
            Se necessário, inicie pelo navegador adequado.
            ''')

        webbrowser.open_new(vpn_url)

    def retry_connection():
        '''Se der erro de conexão, pede pra acessar VPN.'''

        retry = input('Erro de conexão. VPN-Sefaz não está ativado.\n' +
                      'Faça a conexão e aperte qualquer tecla para continuar, ou "S" para sair.\n')
        if retry.lower() == 's':
            print('Programa encerrado pelo usuário.')
            import sys
            sys.exit()

    def pass_login_data():
        '''Pega usuário e senha para passar ao <form> de login'''

        username = driver.find_element_by_xpath(
            '//*[@id="userID2"]').send_keys('blpereira')
        psw = getpass('Insira sua senha: ')
        password = driver.find_element_by_xpath(
            '//*[@id="password"]').send_keys(psw)
        login_button = driver.find_element_by_xpath(
            '//*[@id="ns_Z7_TBK0BB1A0ON6A0I7LI1MU530G5__login"]').click()

    # tenta acessar site do PGSF; se der erro de conexão, abre janela para VPN
    url_login = 'https://portal60.sede.fazenda.sp.gov.br/wps_migrated/portal'
    
    while True:
        try:
            driver.get(url_login)
            break
        # pressuponho que o 'ERR_NAME_NOT_RESOLVED' é sempre pela falta de VPN
        except exceptions.WebDriverException as e:
            if 'ERR_NAME_NOT_RESOLVED' in str(e):
                driver.minimize_window()
                open_vpn()
                retry_connection()

    # pede usuário e senha
    while True:
        pass_login_data()
        # verifica se passou da página de login
        if url_login not in driver.current_url:
            break


def get_osf():
       
    driver.get('https://portal60.sede.fazenda.sp.gov.br/wps_migrated/myportal/!ut/p/b1/04_SjzQ0MzIxtzQ2NjDRj9CPykssy0xPLMnMz0vMAfGjzOIDXUKNjc38DQ38LfwMDYycQn3cLfzdDAzcDfVzoxwVAS6fbJA!/')

    try:
        if WebDriverWait(driver, 4).until(EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, 'iframe'))) is False:
            WebDriverWait(driver, 4).until(
                EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, 'iframe')))
        # tenta duas vezes, pra garantir
    except exceptions.TimeoutException:
        pass

    table = driver.find_element_by_xpath(
        '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr/td/table[3]')
    rows = table.find_elements_by_tag_name('tr')
    rows = rows[3:]  # as três primeiras linhas são cabeçalho

    row = rows[7]
    cells = row.find_elements_by_tag_name('td')

    # choose_osf_button = row.find_element_by_name('idOSF')  # ou a opção abaixo; ambas clicáveis
    choose_osf_button = cells[0].click  # botão clicável
    number_osf, date = cells[1].text.split('\n')
    contribuinte, ie = cells[4].text.split('\n')
    status = cells[6].text

    button_relatar = driver.find_element_by_name('btnRelatarResultado').click
    button_anexar = driver.find_element_by_name('AnexarInformacoes').click
    button_assinar = driver.find_element_by_name('Assinar').click


# exemplo de busca e clique:
'''
for row in rows:
    if '03.0.04110/20-9' in row.text:
        row.find_element_by_name('idOSF').click()

for row in rows:
    if 'Em execução' in row.text:
        print(row.text)
'''
# Para se certificar de que está:
#   - 'aguardando início', para automatizar o colocar em execução
#   - "Em execução", para colocar relatos/enviar arquivos
#   - "em espera de assinatura', para automatizar assinatura ('Em execução')
#   - 'Em espera C Qualidade 1º nível", para ignorar/excluir
# Pode também ser conveniente separar as listas de OSF de acordo com os estados
# que elas se encontram. Uma lista para as 'em execução', outras para as
# 'pedentes de assinatura' etc.


def button(action):
    '''Uso: button(action)
    Botôes existentes:
    relatar resultado, assinar, anexar, confirmar, cancelar, voltar
    '''
    def find_id():
        driver.find_element_by_id(tag_id).click()

    def find_id_swap_case():
        ''' Tenta encontrar o id com a inicial invertida 
        (ex: 'cancelar' --> 'Cancelar'; 'Anexar' --> 'anexar').'''

        tag_id = tag_id[0].swapcase() + tag_id[1:]
        try:
            find_id(tag_id)
        except exceptions.NoSuchElementException:
            pass

    def find_cancelar():
        '''Busca o botão "Cancelar" com o id "Desistir"'''
        try:
            find_id('Desistir')
        except exceptions.NoSuchElementException:
            pass

    actions = {'anexar': 'AnexarInformacoes',
               'assinar': 'Assinar',
               'cancelar': 'Cancelar',
               'confirmar': 'Confirmar',
               'dados': 'dados',
               'desistir': 'Desistir',  # botaão "Cancelar" na tela de anexar informações
               'enviar': 'Enviar',
               'relatar': 'btnRelatarResultado',
               'servico': 'servico',
               'voltar': 'voltar'}

    assert action in actions, 'Ação solicitada não existe/não está mapeada.'

    tag_id = actions[action]

    try:
        find_id()
        return None

    except exceptions.NoSuchElementException:
        # caso não encontre, busca com o caso da inicial invertido;
        # se for 'cancelar', busca também 'Desistir';

        find_id_swap_case()

        if tag_id == 'cancelar':
            find_cancelar()

        # ajuste apenas para exibição
        if tag_id == 'servico':
            tag_id == 'serviço'

        print(f'Erro! Botão {tag_id.title()} indisponível.')
        return 'erro'


# na página de relatar o resultado:
#   - o campo de início estava desabilitado quando acessei; olhar de novo
#   - data_entrega_osf_id = 'dtcEntregaOSF'
#   - data_termino_osf_id = 'dtcTermino'
#

#data_entrega = driver.find_element_by_id('dtcEntregaOSF').send_keys
#data_termino = driver.find_element_by_id('dtcTermino').send_keys


def escreve_observacoes(txt):
    try:
        driver.find_element_by_id('txtBxObservacoes').clear()
        driver.find_element_by_id('txtBxObservacoes').send_keys(txt)

        # TODO: abrir messagebox pedindo para confirmar inserção
        # confirma = driver.find_element_by_id('Confirmar').click
        # cancela = driver.find_element_by_id('Cancelar').click
        # volta = driver.find_element_by_id('Voltar').click

        # confirmação de que deu certo:
        # texto 'O resultado da OSF foi relatado com sucesso.' in driver.page_source
        return 0
    except exceptions.NoSuchElementException:
        print('Caixa de texto não encontrada.')
        return 1


# lidando com alertas JS
Alert = webdriver.common.alert.Alert
# Alert(driver).text: pega texto
# Alert(driver).dismiss()
# Alert(driver).accept()
# exceção: selenium.common.exceptions.NoAlertPresentException

login()
