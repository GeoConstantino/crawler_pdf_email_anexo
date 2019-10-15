import fitz
import glob
import pandas as pd


def parse_email(pdf_name):
    lista_metadados = list()
    pdf = fitz.open(pdf_name)

    for page in range(pdf.pageCount):
        page_read = pdf.loadPage(page)
        page_text = page_read.getText("text").split('\n')
        metadados = get_metadata_email(page_text)
        if metadados:
            #import ipdb; ipdb.set_trace()
            metadados['Localização no PDF'] = 'Página {n_pagina} do arquivo {n_arquivo}'.format(n_pagina = page+1, n_arquivo = pdf_name.split('/')[1])
            metadados['Data Criação do PDF'] = format_data(pdf.metadata['creationDate'])
            metadados['Data Alteração do PDF'] = format_data(pdf.metadata['modDate'])
            lista_metadados.append(metadados)
       
    return lista_metadados


def format_data(raw):
    
    ano = raw[2:6]
    mes = raw[6:8]
    dia = raw[8:10]
    hora = raw[10:12]
    minuto = raw[12:14]
    segundo = raw[14:16]

    data = '{dd}/{mm}/{aaaa} {h}:{m}:{s}'.format(
        dd=dia, mm=mes, aaaa=ano, h=hora, m=minuto, s=segundo
        )

    return data

    
def get_metadata_email(page_text):

    metadados = dict()
    keywords = ['De:', 'Enviado em:' , 'Para:', 'Cc:', 'Assunto:', 'Anexos:', 'Prioridade:']

    for i, item in enumerate(page_text):
        if item in keywords:
            metadados[item] = str(page_text[i+1]).strip().replace("'","")
    
    return metadados


def get_data_envio(data):

    try:
        _data = data.split(',')[1].strip()
        dia, mes, ano, hora = _data.replace(' de ', ' ').strip().split(' ')
        data_fim = '{}/{}/{} {}'.format(dia, mes_para_numero(mes), ano, hora)
        return data_fim
    except:
        #import ipdb; ipdb.set_trace()
        return data
    

def mes_para_numero(mes):

    num = { 'janeiro':'01', 
            'fevereiro':'02', 
            'março':'03',
            'abril':'04', 
            'maio':'05', 
            'junho':'06', 
            'julho':'07', 
            'agosto':'08', 
            'setembro':'09', 
            'outubro':'10', 
            'novembro':'11', 
            'dezembro':'12'}

    return num[mes]


if __name__ == "__main__":

    df = pd.DataFrame()
    pdfs = glob.glob('emails/*.pdf')

    for pdf_name in pdfs:

        df = df.append(parse_email(pdf_name))
        #import ipdb; ipdb.set_trace()

    df = df[['De:', 'Enviado em:', 'Para:', 'Cc:', 'Assunto:', 'Anexos:', 'Prioridade:','Localização no PDF','Data Criação do PDF','Data Alteração do PDF']]

    df['Enviado em (formatado):'] = df['Enviado em:'].apply(lambda x: get_data_envio(x))
    
    writer = pd.ExcelWriter('Leitura_PDF.xlsx', engine='xlsxwriter', datetime_format='DD/MM/YYYY HH:MM')
    
    df.to_excel(writer, index=False, sheet_name='Planilha1')

    workbook = writer.book
    worksheet = writer.sheets['Planilha1']
    worksheet.set_column('A:K', 30)
    writer.save()
    writer.close()

