import os
import re
from tkinter import messagebox


def parse_osf(txt):
    '''Extrai dados da OSF.
    Retorna dicionário com os dados.
    '''
    
    osf_regex = re.compile(r'Número da OSF:\n+(.*)')
    nome_regex = re.compile(r'Razão Social:\n+(.*)')
    logradouro_regex = re.compile(r'Endereço:\n+(.*)')
    número_regex = re.compile(r'N.:\n(.*)')
    bairro_regex = re.compile(r'Bairro:\n+(.*)')
    cep_regex = re.compile(r'CEP:\n+(.*)')
    cidade_regex = re.compile('Município:\n+(.*)')
    ie_regex = re.compile(r'Inscrição Estadual:\n+(.*)')
    cnpj_regex = re.compile(r'CNPJ/CPF:\n+(.*)')

    osf = dict()

    osf['OSF'] = osf_regex.search(txt)[1].strip()
    osf['Nome'] = nome_regex.search(txt)[1].strip()
    osf['IE'] = ie_regex.search(txt)[1].strip()
    osf['CNPJ'] = cnpj_regex.search(txt)[1].strip()
    osf['Logradouro'] = logradouro_regex.search(txt)[1].strip()
    osf['Número'] = número_regex.search(txt)[1].strip()
    osf['Bairro'] = bairro_regex.search(txt)[1].strip()
    osf['CEP'] = cep_regex.search(txt)[1].replace('.', '').strip()
    osf['Cidade'] = cidade_regex.search(txt)[1].strip()
    osf['Cidade'] = acentuar_cidades(osf['Cidade'])
    osf['Endereço'] = f"{osf['Logradouro']}, {osf['Número']} - CEP {osf['CEP']} - {osf['Cidade']}/SP"

    osf['cabecalho'] = f'''CONTRIBUINTE: {osf['Nome']}
    IE: {osf['IE']}
    CNPJ: {osf['CNPJ']}
    ENDEREÇO: {osf['Endereço']}
    REF: OSF n. {osf}
    '''.replace('  ', ' ')

    return osf


def is_osf(txt, filename):
    '''Verifica se o texto passado é uma OSF. 
    Retorna True ou False'''
    
    if 'ORDEM DE SERVIÇO FISCAL - OSF' not in txt:
        messagebox.showinfo(
            'Erro', f'Arquivo {filename} não é uma OSF ou o arquivo está corrompido.')
        print(f'Erro. Arquivo {filename} não é uma OSF.')
    
        return False
    
    return True
    

def acentuar_cidades(txt):
    '''Coloca acento no nome das cidades'''

    dic_correções = ({
                     'AO ': 'ÃO ',
                     'REDENCAO': 'REDENÇÃO',
                     'CACAPAVA': 'CAÇAPAVA',
                     'JORDAO': 'JORDÃO',
                     'GUARATINGUETA': 'GUARATINGUETÁ',
                     'IGARATA': 'IGARATÁ',
                     'JACAREI': 'JACAREÍ',
                     'SAPUCAI': 'SAPUCAÍ',
                     'TAUBATE': 'TAUBATÉ',
                     'TREMEMBE': 'TREMEMBÉ'
                     })

    for k, v in dic_correções.items():
        txt = txt.replace(k, v)

    return txt
