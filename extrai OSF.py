import os
import io
import pyperclip
import sys

from pathlib import Path

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams

from selenium.common import exceptions

from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilename

import browser
import parse_osf


def choose_file():

    # askopenfilename importado de tkinter.fileadialog
    file = askopenfilename(initialdir=working_dir,
                           filetypes=[("arquivos de OSF", "OSF*.pdf"),
                                      ("arquivos PDF", "*.pdf"),
                                      ("todos arquivos", "*.*")])
    dock.update()

    # return Path('C:/Users/bruno/OneDrive/Arquivos trabalho/Nos Conformes/1.0 - OrdemServicoFiscal.pdf')
    return Path(file)


def get_text(file):
    '''Extrai a primeira página do PDF da OSF.
    Retorna o texto extraído.'''

    # setup do extrator
    retstr = io.StringIO()
    device = TextConverter(PDFResourceManager(), retstr, laparams=LAParams())
    interpreter = PDFPageInterpreter(PDFResourceManager(), device)

    with open(file, 'rb') as pdf:
        page = PDFPage.get_pages(pdf)  # cria um generator com para as pgs

        page = next(page)  # pega apenas primeira p. do generator
        interpreter.process_page(page)
        text = retstr.getvalue()

        device.close()
        retstr.close()

    return text


def busca_google_maps():
    this.driver.get(
        f"https://www.google.com/maps/place/{osf['Logradouro']} {osf['Número']}, {osf['Cidade']}")


class Botao:
    '''
    Cria a classe dos botões, recebendo o nome de apresentação e a variável
    a ser utilizada. 
    Ao clicar, copia variável para área de transferência.

    Parâmetros:
        txt = nome do Botao
        var = informação a ser copiada ao clicar no Botao
    '''

    def __init__(self, campo, dados):
        display = dados[:34] + (dados[34:] and '...')
        self.campo = campo
        self.dados = dados
        self.text = StringVar()
        self.text.set(f'{self.campo}: {display}')
        self.create = Button(dock, textvariable=self.text, width=40,
                             command=lambda: pyperclip.copy(self.dados),
                             padx=10, anchor='w')
        self.create.pack()

    def atualiza_texto(self, osf):
        '''Atualiza o texto do GUI buscando o valor atual da chave no dicionário'''
        self.dados = osf[self.campo]
        display = self.dados[:34] + (self.dados[34:] and '...')
        self.text.set(f'{self.campo}: {display}')
        #print(f'{self.campo}: {display}')
        # self.text.set(display)


def widget():
    '''Cria o widget com os botões'''
    def atualiza():
        this.osf = extract()
        if this.osf is None:
            return None

        for botao in botoes:
            botao.atualiza_texto(osf)

    topo = Label(dock, text='Copiar para área de transferência')
    topo["font"] = ("Verdana", "10", 'bold')
    topo.pack()

    # cria botões iterando sobre o dicionário OSF
    botoes = []
    for k, v in osf.items():
        if k == 'cabecalho':
            continue
        botoes.append(Botao(k, v))

    Label(dock).pack()  # espaço em branco

    meio = Label(dock, text='Gerar documentos no SigaDoc')
    meio["font"] = ("Verdana", "10", 'bold')
    meio.pack()
    Button(dock, text='Expediente SigaDoc',
           command=cria_expediente, width=40).pack()
    Button(dock, text='Formulário 2.05P',
           command=formulario_205p, width=40).pack()
    Button(dock, text='Fotografias', command=junta_fotos, width=40).pack()
    Label(dock).pack()  # espaço em branco

    Button(dock, text='Abrir outra OSF', command=atualiza, width=20).pack()
    Button(dock, text='Sair', command=dock.destroy, width=20).pack()

    dock.mainloop()


def verifica_sessao():

    # Se não tiver driver instanciado, cria um
    if this.driver is None:
        launch_driver()

    # Se o driver não estiver respondendo, abre outro
    if browser.driver_ativo(this.driver) == False:
        launch_driver()

    while not this.session:
        this.session = acessa_sistema()

    return True


def launch_driver():
    '''Inicia sessão do Selenium WebDriver.'''

    this.driver = browser.start(
        path='C:/Users/bruno/OneDrive/Arquivos trabalho/Nos Conformes', download_pdf=False)
    this.session = False


def acessa_sistema():
    '''Faz login no Sem Papel.
    Retorna True se a sessão foi iniciada.'''

    this.session = True
    verifica_sessao()

    driver = this.driver

    # Acesso ao sistema

    driver.get("https://www.documentos.spsempapel.sp.gov.br/siga/public/app/login?cont=https%3A%2F%2Fwww.documentos.spsempapel.sp.gov.br%2Fsiga%2Fapp%2Fprincipal")
    driver.find_element("id", "username").send_keys("SFP31451")
    driver.find_element("id", "password").send_keys("fEni1$28")
    driver.find_element_by_css_selector(".btn-lg").click()

    while True:
        if 'login' in driver.title.lower():
            continuar = messagebox.askokcancel('Erro ao entrar',
                                               'Faça o login manualmente e clique em OK, para\n' +
                                               'continuar, ou clique em "Cancelar" para sair')
            if continuar is False:
                this.session = False
                sys.exit()
        else:
            break

    # se estiver substituindo outra pessoa, encerra substituição
    try:
        button = driver.find_element(
            'xpath', '/html/body/div[1]/div[1]/div[2]/div[2]/button')
        if button.text == 'Finalizar':
            button.click()
    except:
        pass

    return True


def cria_expediente():
    '''Cria o expediente de Não Localização no Sem Papel, com preenchimento automático
    das informações.'''

    verifica_sessao()
    # Verifica se webdriver está funcionando
    # if verifica_sessao() == False:
    #     launch_driver()
    #     this.session = acessa_sistema()

    # Verifica se está logado
    # while not this.session:
    #     this.session = acessa_sistema()

    # Vai para página de edição e preenche dados
    driver = this.driver

    driver.get(
        'https://www.documentos.spsempapel.sp.gov.br/sigaex/app/expediente/doc/editar')
    driver.find_element("id", 'dropdownMenuButton').click()
    driver.find_element('xpath',
                        "//a[contains(text(),'Expediente de verificação fiscal')]").click()
    driver.find_element_by('name', "complemento_do_assunto").send_keys(
        'Não localização - Programa Nos Conformes')
    driver.find_element_by('name', "interessado_nome").send_keys(osf['Nome'])
    driver.find_element_by('name', "interessado_cnpj").send_keys(osf['CNPJ'])
    driver.find_element_by('name', "interessado_ie").send_keys(osf['IE'])
    driver.find_element_by('name',
                           "interessado_logradouro").send_keys(osf['Logradouro'])
    driver.find_element_by(
        'name', "interessado_numero").send_keys(osf['Número'])
    # driver.find_element_by('name', "interessado_endereco_complemento").send_keys("")
    # driver.find_element_by('name', "interessado_bairro").send_keys("")
    driver.find_element_by('name',
                           "interessado_municipio").send_keys(osf['Cidade'])
    driver.find_element_by('name', "interessado_municipio_uf").click()
    driver.find_element('xpath', "//option[@value='SP']").click()
    driver.find_element_by('name', "interessado_cep").send_keys(osf['CEP'])
    driver.find_element_by('name', "ver_doc_pdf").click()
    # driver.find_element("id","btnGravar").click()  # Botao de confirmação


def está_na_tela_processo():
    """Verifica se driver está aberto e se está na tela do processo.
    """
    # Verifica se webdriver está funcionando
    # if verifica_sessao() == False:
    #     launch_driver()
    #     this.session = acessa_sistema()

    # # Verifica se está logado
    # while not this.session:
    #     this.session = acessa_sistema()
    verifica_sessao()

    driver = this.driver

    # adicionar EXP ou PRC
    url_base_processo = 'https://www.documentos.spsempapel.sp.gov.br/sigaex/app/expediente/doc/exibir?sigla='

    while ((url_base_processo not in driver.current_url) or
           (osf['Nome'] not in driver.find_element_by_tag_name('body').text)):
        continuar = messagebox.askokcancel(
            "Execução suspensa", "Navegue até a página do expediente \ne clique 'OK' para continuar")
        if continuar is False:
            return False
    return True


def formulario_205p():
    'Preenche o Demonstrativo 2.05P com os dados obtidos.'
    # A parte da assinatura e da efetivaçao da criação ainda está manual.

    driver = this.driver

    if está_na_tela_processo():
        driver.find_element_by_link_text('Incluir Documento').click()
        driver.find_element("id", 'dropdownMenuButton').click()
        driver.find_element('xpath',
                            "//a[contains(.,'Roteiro 2.05-P - Não Localização - Programa Nos Conformes')]").click()
        driver.find_element_by('name', 'complemento_do_assunto').send_keys(
            'Demonstrativo de Não Localização - 2.05P - Nos Conformes')
        driver.find_element_by(
            'name', 'interessado_nome').send_keys(osf['Nome'])
        driver.find_element_by(
            'name', 'interessado_cnpj').send_keys(osf['CNPJ'])
        driver.find_element_by('name', 'interessado_ie').send_keys(osf['IE'])
        driver.find_element_by('name',
                               'interessado_logradouro').send_keys(osf['Logradouro'])
        driver.find_element_by('name',
                               'interessado_numero').send_keys(osf['Número'])
        driver.find_element_by('name',
                               'interessado_bairro').send_keys(osf['Bairro'])
        driver.find_element_by('name',
                               "interessado_municipio").send_keys(osf['Cidade'])
        driver.find_element_by('name', "interessado_municipio_uf").click()
        driver.find_element('xpath', "//option[@value='SP']").click()
        driver.find_element_by('name', 'interessado_cep').send_keys(osf['CEP'])
        driver.find_element_by('name', 'numero_osf').send_keys(osf['OSF'])
        messagebox.showinfo('Ação do usuário',
                            'Insira a data da visita e preencha o relato.')

        # Não estou conseguindo preencher a parte do texto.
        # Considerar criar o texto no script e mandar para área de transf.
        # driver.find_element_by_css_selector('html').click()
        # driver.find_element_by_css_selector('html').send_keys('Relato aqui')
    else:
        messagebox.showinfo(
            'Fim da rotina', 'Execução interrompida pelo usuário.')


def junta_fotos(verifica=True):
    '''Acessa "Incluir documento" e preenche campos para juntada das fotografias.
    Por segurança, verifica se a URL atual é a do documento que foi gerado ao
    executar form_205P().'''

    driver = this.driver

    if está_na_tela_processo():
        driver.find_element_by_link_text('Incluir Documento').click()
        # Clica em "Selecione o modelo"
        driver.find_element('xpath',
                            "//button[@id='dropdownMenuButton']/span").click()
        driver.find_element_by_link_text(
            "Documento CapturadoDocumento Capturado").click()
        driver.find_element_by('name',
                               'Assunto').send_keys("Fotografia do local")

        # Clica em "Tipo do Documento"
        driver.find_element_by('name', "especie").click()

        for i in range(3):
            try:
                driver.find_element('xpath',
                                    "//option[contains(.,'Outros')]").click()
                driver.implicitly_wait(2)
            except exceptions.NoSuchElementException:
                pass

        # Clica em "Tipo de Conferência"
        driver.find_element_by('name', "conferencia").click()
        driver.find_element('xpath',
                            "//option[@value='Documento Original']").click()

        # Preenche "Descrição"
        driver.find_element_by('name', "outros").send_keys("Fotografia")
    else:
        messagebox.showinfo(message='Execução interrompida pelo usuário.')


def clear_osf():
    for key in osf:
        osf[key] = ''


def extract():

    while True:
        file = choose_file()

        # se não selecionou nada, retorna None
        if file.is_file() is False:
            return None

        txt = get_text(file)

        if parse_osf.is_osf(txt, file.name) is True:
            clear_osf()
            break
    # TODO: tratar da situação de o usuário não abrir nenhum arquivo.
    # Manter os dados que

    this.osf = parse_osf.parse_osf(txt)
    print(osf['cabecalho'])

    return osf


# setup
working_dir = Path(
    'C:/Users/bruno/OneDrive/Arquivos trabalho/Nos conformes')
os.chdir(working_dir)

# Interface gráfica
dock = Tk()
dock.title('Extrai OSF')

# cria pointers para driver e session
this = sys.modules[__name__]
session = bool()
driver = None
osf = {}
# extrai dados

osf = extract()
if osf is None:
    sys.exit()

widget()
