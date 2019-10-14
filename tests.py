from unittest import TestCase
import fitz

from main import get_metadata_email, format_data, get_data_envio


class ReadEmail(TestCase):

    def test_read_email_1(self):
        email = ['De:', 'Maria Aparecida Valenca Bezerra de Menezes', 'Enviado em:', 'sexta-feira, 27 de outubro de 2006 09:57', 'Para:', 'Leandro Frota Duarte; Ricardo Simonsen', 'Assunto:', 'Relatório Berj', 'Anexos:', 'Relatório Final.doc', 'Prioridade:', 'Alta', 'Por favor utilizem o arquivo anexo para fazer as modificações. ', ' ', 'Mª Aparecida Valença ', 'FGV Projetos ', 'Vice-diretoria de Projetos ', 'Tel.: (21) 2559-5624 ', '']

        meta = get_metadata_email(email)

        self.assertEqual(meta['De:'], 'Maria Aparecida Valenca Bezerra de Menezes')
        self.assertEqual(meta['Enviado em:'], 'sexta-feira, 27 de outubro de 2006 09:57')
        self.assertEqual(meta['Para:'], ['Leandro Frota Duarte','Ricardo Simonsen'])
        self.assertEqual(meta['Assunto:'], 'Relatório Berj')
        self.assertEqual(meta['Anexos:'], ['Relatório Final.doc'])
        self.assertEqual(meta['Prioridade:'], 'Alta')
        #self.assertEqual(meta['DataCriaçãoArquivoPDF'], '')

    def test_data_criacao_email(self):

        pdf_name = "teste_pdf/1.pdf"
        pdf = fitz.open(pdf_name)
        
        self.assertEqual(format_data(pdf.metadata['creationDate']),'03/06/2019 10:24:38')

    def test_data_envio_email(self):

        data = 'sexta-feira, 27 de outubro de 2006 09:57'

        self.assertEqual(get_data_envio(data), '27/06/2019 09:57:00')


