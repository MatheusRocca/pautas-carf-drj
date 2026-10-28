# -*- coding: utf-8 -*-
"""
Testes do extrator de pautas CARF/DRJ, usando XMLs reais do INLabs em
tests/fixtures/.

Rodar (da pasta do projeto):
    python3 -m unittest discover tests -v

O teste de "snapshot" compara a extração de cada XML com o resultado
conferido e salvo em tests/fixtures/esperado.json. Se uma mudança no
script alterar a extração de propósito, confira a diferença e regrave com:
    ATUALIZAR_ESPERADO=1 python3 -m unittest discover tests
"""

import contextlib
import glob
import importlib.util
import io
import json
import os
import tempfile
import unittest
from unittest import mock
from datetime import date, datetime

from openpyxl import Workbook, load_workbook

PASTA_TESTES = os.path.dirname(os.path.abspath(__file__))
PASTA_FIXTURES = os.path.join(PASTA_TESTES, 'fixtures')
ARQ_ESPERADO = os.path.join(PASTA_FIXTURES, 'esperado.json')
SCRIPT = os.path.join(PASTA_TESTES, '..', 'pautas_carf.py')

# o nome do script tem espaços, então não dá pra fazer "import" normal
_spec = importlib.util.spec_from_file_location('pauta', SCRIPT)
pauta = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(pauta)


def _parse(nome_arquivo):
    with contextlib.redirect_stdout(io.StringIO()):
        return pauta.parse_xml_pauta(os.path.join(PASTA_FIXTURES, nome_arquivo))


def _fixtures_pauta():
    return sorted(os.path.basename(f) for f in glob.glob(os.path.join(PASTA_FIXTURES, '515_*.xml')))


# ─────────────────────────────────────────────
# Extração dos XMLs reais
# ─────────────────────────────────────────────

class TestExtracaoXmlReais(unittest.TestCase):

    def test_snapshot_de_todas_as_pautas(self):
        atual = {nome: _parse(nome) for nome in _fixtures_pauta()}

        if os.environ.get('ATUALIZAR_ESPERADO') or not os.path.exists(ARQ_ESPERADO):
            with open(ARQ_ESPERADO, 'w', encoding='utf-8') as fp:
                json.dump(atual, fp, ensure_ascii=False, indent=1)
            self.skipTest(f'esperado.json regravado com {len(atual)} XMLs — confira antes de confiar')

        with open(ARQ_ESPERADO, encoding='utf-8') as fp:
            esperado = json.load(fp)
        self.assertEqual(sorted(atual), sorted(esperado), 'conjunto de XMLs mudou')
        for nome in esperado:
            with self.subTest(xml=nome):
                self.assertEqual(atual[nome], esperado[nome])

    def test_nenhum_processo_perdido(self):
        """Cada 'N - Processo nº' do texto vira exatamente uma linha."""
        for nome in _fixtures_pauta():
            with self.subTest(xml=nome):
                import xml.etree.ElementTree as ET
                art = ET.parse(os.path.join(PASTA_FIXTURES, nome)).getroot().find('article')
                texto = pauta.html_para_texto(art.findtext('.//Texto') or '')
                esperados = len(pauta.PAT_PROCESSO_MARCADOR.findall(texto))
                self.assertEqual(len(_parse(nome)), esperados)

    def test_campos_obrigatorios_preenchidos(self):
        for nome in _fixtures_pauta():
            for linha in _parse(nome):
                with self.subTest(xml=nome, processo=linha['Número do Processo']):
                    self.assertRegex(linha['Número do Processo'], r'^\d{5}\.\d{6}/\d{4}-\d{2}$')
                    for campo in ('Recorrente', 'Interessado', 'Relator(a)', 'Data de Julgamento', 'Horário'):
                        self.assertTrue(linha[campo], f'{campo} vazio')

    def test_xml_que_nao_e_pauta_e_ignorado(self):
        self.assertEqual(_parse('nao_pauta.xml'), [])

    def test_pautas_de_outros_orgaos_nao_geram_linhas(self):
        # CRSFN, MPT e CNAS também publicam "Pauta", mas em outro formato
        for nome in ('515_20260925_24413541.xml', '515_20260925_24413975.xml',
                     '515_20260928_24416795.xml', '515_20260928_24419565.xml'):
            with self.subTest(xml=nome):
                self.assertEqual(_parse(nome), [])

    def test_assinatura_em_maiusculas_nao_gruda_no_ultimo_processo(self):
        # Nesses XMLs o presidente da turma assina em MAIÚSCULAS
        # (<p class="assina">MAURICIO SOARES RODRIGUES</p>).
        for nome in ('515_20260925_24408020.xml', '515_20260928_24413870.xml',
                     '515_20260928_24419830.xml', '515_20260928_24419831.xml'):
            with self.subTest(xml=nome):
                self.assertEqual(_parse(nome)[-1]['Interessado'], 'FAZENDA NACIONAL')

    def test_relator_muda_no_meio_da_sessao(self):
        linhas = _parse('515_20260928_24413870.xml')
        relatores = [(l['Número do Processo'], l['Relator(a)']) for l in linhas]
        self.assertEqual(relatores[2], ('75300.119721/2021-91', 'ORLANDO SANTIAGO DA COSTA JUNIOR'))
        self.assertEqual(relatores[3], ('25579.415521/2021-25', 'FERNANDA AMARAL SARNO DE ALMEIDA'))

    def test_exemplo_carf(self):
        linha = _parse('515_20260925_24410326.xml')[0]
        self.assertEqual(linha['Órgão'], 'Ministério da Fazenda / Conselho Administrativo de Recursos Fiscais')
        self.assertEqual(linha['Número do Processo'], '10278.681447/2018-65')
        self.assertEqual(linha['Recorrente'], 'BACAMON SATRAQUI MILUNE')
        self.assertEqual(linha['Data de Julgamento'], '5 de Outubro de 2026')
        self.assertEqual(linha['Horário'], '09:00')
        self.assertEqual(linha['Data Publicação DOU'], '25/09/2026')


# ─────────────────────────────────────────────
# Casos sintéticos de parsing
# ─────────────────────────────────────────────

def _xml_pauta(categoria, texto_html):
    from xml.sax.saxutils import escape
    return ('<?xml version="1.0" encoding="UTF-8"?><xml><article artType="Pauta" '
            f'artCategory="{escape(categoria)}" pubDate="25/09/2026" pdfPage="http://exemplo/pagina">'
            f'<body><Texto>{escape(texto_html)}</Texto></body></article></xml>')


class TestAvisosDeParsing(unittest.TestCase):

    def parse(self, categoria, texto_html):
        with tempfile.TemporaryDirectory() as tmp:
            caminho = os.path.join(tmp, 'pauta.xml')
            with open(caminho, 'w', encoding='utf-8') as fp:
                fp.write(_xml_pauta(categoria, texto_html))
            avisos = []
            with contextlib.redirect_stdout(io.StringIO()):
                linhas = pauta.parse_xml_pauta(caminho, avisos)
        return linhas, avisos

    def test_pauta_carf_sem_processos_avisa(self):
        # simula mudança de formato: o texto não tem mais "1 - Processo nº"
        linhas, avisos = self.parse('Ministério da Fazenda/Conselho Administrativo de Recursos Fiscais',
                                    '<p>Processo 43106.903311/2006-41, Recorrente EMPRESA X</p>')
        self.assertEqual(linhas, [])
        self.assertEqual(len(avisos), 1)
        self.assertIn('formato pode ter mudado', avisos[0])
        self.assertIn('http://exemplo/pagina', avisos[0])

    def test_pauta_drj_sem_processos_avisa(self):
        _, avisos = self.parse('Ministério da Fazenda/Secretaria Especial da Receita Federal do Brasil/'
                               'Delegacia de Julgamento da Receita Federal do Brasil 04/Sexta Turma', '<p>nada</p>')
        self.assertEqual(len(avisos), 1)

    def test_pauta_de_outro_orgao_sem_processos_nao_avisa(self):
        _, avisos = self.parse('Ministério da Fazenda/Conselho de Recursos do Sistema Financeiro Nacional',
                               '<p>nada</p>')
        self.assertEqual(avisos, [])

    def test_processo_perdido_avisa(self):
        # o 2º processo não tem rótulos de parte e some na extração
        html = ('<p>1 - Processo nº: 43106.903311/2006-41 - Recorrente: EMPRESA A e Interessado: FAZENDA NACIONAL</p>'
                '<p>2 - Processo nº 90790.306738/2006-41</p>')
        with mock.patch.object(pauta, 'extrair_processos', return_value=[]):
            _, avisos = self.parse('Ministério da Fazenda/Conselho Administrativo de Recursos Fiscais', html)
        self.assertEqual(len(avisos), 1)
        self.assertIn('2 processo(s) numerado(s) no texto, mas 0 extraído(s)', avisos[0])

    def test_pautas_reais_nao_geram_avisos(self):
        for nome in _fixtures_pauta():
            avisos = []
            with contextlib.redirect_stdout(io.StringIO()):
                pauta.parse_xml_pauta(os.path.join(PASTA_FIXTURES, nome), avisos)
            with self.subTest(xml=nome):
                self.assertEqual(avisos, [])


class TestParsingSintetico(unittest.TestCase):

    def test_html_ignora_assinatura_e_cargo(self):
        html = ('<p>1 - Processo nº: 1 - Recorrente: X e Interessado: FAZENDA NACIONAL</p>'
                '<p class="assina">FULANO DE TAL</p><p class="cargo">Presidente da Turma</p>')
        texto = pauta.html_para_texto(html)
        self.assertNotIn('FULANO', texto)
        self.assertNotIn('Presidente', texto)

    def test_varias_partes_e_rotulos_variados(self):
        texto = ('1 - Processo nº: 10880.000001/2020-11 - Embargante: EMPRESA A LTDA e '
                 'Embargado: FAZENDA NACIONAL e Responsavel: JOAO DA SILVA')
        proc, = pauta.extrair_processos(texto)
        self.assertEqual(proc['recorrente'], 'EMPRESA A LTDA')
        self.assertEqual(proc['interessado'], 'FAZENDA NACIONAL')
        self.assertEqual(proc['outras_partes'], 'Responsavel: JOAO DA SILVA')


# ─────────────────────────────────────────────
# Checagem de clientes (dados sintéticos — a lista real não entra nos testes)
# ─────────────────────────────────────────────

def _cliente(nome):
    return {'nome': nome, 'doc': '', 'nome_base': pauta.nome_base(pauta.normalizar_nome(nome))}


class TestChecagemClientes(unittest.TestCase):

    def setUp(self):
        self.clientes = [_cliente('Alfa Brasil Nordeste Ltda'), _cliente('Beta Química S/A'),
                         _cliente('ABC')]
        self.processos = {
            '43106903311200641': ['Alfa Brasil Nordeste Ltda'],   # número + nome
            '48356420873202111': ['Outra Empresa Ltda'],         # número com nome diferente
            '99999999999999999': [],                             # número sem nome
        }

    def checar(self, numero, recorrente, interessado='FAZENDA NACIONAL'):
        return pauta.checar_cliente(numero, recorrente, interessado, '', self.clientes, self.processos)

    def test_numero_e_nome_confirmados(self):
        nome, tipo, _ = self.checar('43106.903311/2006-41', 'ALFA BRASIL NORDESTE LTDA.')
        self.assertEqual(nome, 'Alfa Brasil Nordeste Ltda')
        self.assertIn('confirmado', tipo)

    def test_numero_com_nome_diferente(self):
        nome, tipo, _ = self.checar('48356.420873/2021-11', 'EMPRESA QUALQUER LTDA')
        self.assertEqual(nome, 'Outra Empresa Ltda')
        self.assertEqual(tipo, 'Número do processo (confira o nome)')

    def test_numero_sem_nome_cadastrado(self):
        nome, tipo, _ = self.checar('99999.999999/9999-99', 'EMPRESA QUALQUER LTDA')
        self.assertEqual(nome, 'EMPRESA QUALQUER LTDA / FAZENDA NACIONAL')
        self.assertIn('sem nome cadastrado', tipo)

    def test_so_pelo_nome_ignorando_acento_e_sufixo(self):
        nome, tipo, _ = self.checar('11111.111111/2020-11', 'BETA QUIMICA S/A')
        self.assertEqual(nome, 'Beta Química S/A')
        self.assertTrue(tipo.startswith('Nome'))

    def test_sem_match(self):
        self.assertEqual(self.checar('11111.111111/2020-11', 'NINGUEM LTDA'), ('', None, ''))

    def test_nome_curto_demais_nao_bate(self):
        self.assertEqual(self.checar('11111.111111/2020-11', 'ABC COMERCIO LTDA'), ('', None, ''))

    def test_nome_parcial_nao_bate_no_meio_da_palavra(self):
        self.clientes = [_cliente('Brasil Norte')]
        self.assertEqual(self.checar('11111.111111/2020-11', 'BRASILNORTEX LTDA'), ('', None, ''))


class TestCarregarClientes(unittest.TestCase):

    def test_planilha_com_celulas_mescladas(self):
        # mesmo formato da lista real: nome + processo, com o processo
        # mesclado valendo para várias empresas
        wb = Workbook()
        ws = wb.active
        ws.append(['Partes e empresas monitoradas', 'Processo - Controle interno'])
        ws.append(['Empresa A Ltda', '43106.903311/2006-41'])
        ws.append(['Empresa B S/A', None])
        ws.append([None, '48356.420873/2021-11'])
        ws.merge_cells('B2:B3')
        with tempfile.TemporaryDirectory() as tmp:
            caminho = os.path.join(tmp, 'clientes.xlsx')
            wb.save(caminho)
            clientes, processos = pauta.carregar_clientes(caminho)

        self.assertEqual([c['nome'] for c in clientes], ['Empresa A Ltda', 'Empresa B S/A'])
        self.assertEqual(processos['43106903311200641'], ['Empresa A Ltda', 'Empresa B S/A'])
        self.assertEqual(processos['48356420873202111'], [])

    def test_sem_coluna_de_nome_da_erro(self):
        with tempfile.TemporaryDirectory() as tmp:
            caminho = os.path.join(tmp, 'clientes.csv')
            with open(caminho, 'w', encoding='utf-8') as fp:
                fp.write('Coluna X,Processo\nfoo,123\n')
            with self.assertRaises(ValueError):
                pauta.carregar_clientes(caminho)


# ─────────────────────────────────────────────
# Geração do Excel
# ─────────────────────────────────────────────

class TestTextoParaData(unittest.TestCase):

    def test_formatos_reconhecidos(self):
        self.assertEqual(pauta.texto_para_data('7 de Outubro de 2026'), date(2026, 10, 7))
        self.assertEqual(pauta.texto_para_data('1 de março de 2027'), date(2027, 3, 1))
        self.assertEqual(pauta.texto_para_data('25/09/2026'), date(2026, 9, 25))

    def test_formato_desconhecido_mantem_texto(self):
        for texto in ('', 'amanhã', '31 de Fevereiro de 2026', '7 de Outubrox de 2026'):
            with self.subTest(texto=texto):
                self.assertEqual(pauta.texto_para_data(texto), texto)


class TestGerarXlsx(unittest.TestCase):

    def setUp(self):
        self.linhas = _parse('515_20260925_24408020.xml')
        self.linhas[0].update({'Cliente Identificado': 'Alfa', 'Tipo de Match': 'Nome (confira)', 'Por que bateu': 'x'})
        for l in self.linhas[1:]:
            l.update({'Cliente Identificado': '', 'Tipo de Match': '', 'Por que bateu': ''})

    def _gerar(self, com_clientes=True):
        with tempfile.TemporaryDirectory() as tmp:
            caminho = os.path.join(tmp, 'saida.xlsx')
            with contextlib.redirect_stdout(io.StringIO()):
                pauta.gerar_xlsx(self.linhas, caminho, com_clientes=com_clientes)
            return load_workbook(caminho)

    def _celula(self, ws, coluna, linha=2):
        return ws.cell(row=linha, column=[c.value for c in ws[1]].index(coluna) + 1)

    def test_gera_abas_e_colunas(self):
        wb = self._gerar()
        self.assertEqual(wb.sheetnames, ['Pautas', 'Clientes na Pauta', 'Legenda'])
        cabecalho = [c.value for c in wb['Pautas'][1]]
        self.assertEqual(cabecalho, pauta.COLUNAS_BASE + pauta.COLUNAS_CLIENTE)
        self.assertEqual(wb['Pautas'].cell(row=2, column=3).value, '43106.903311/2006-41')
        self.assertEqual(wb['Clientes na Pauta'].cell(row=2, column=3).value, '43106.903311/2006-41')
        self.assertIsNone(wb['Clientes na Pauta'].cell(row=3, column=3).value)

    def test_sem_clientes_nao_tem_legenda(self):
        self.assertEqual(self._gerar(com_clientes=False).sheetnames, ['Pautas'])

    def test_datas_sao_datas_de_verdade(self):
        ws = self._gerar()['Pautas']
        for coluna, esperado in (('Data de Julgamento', datetime(2026, 10, 8)),
                                 ('Data Publicação DOU', datetime(2026, 9, 25))):
            with self.subTest(coluna=coluna):
                c = self._celula(ws, coluna)
                self.assertEqual(c.value, esperado)
                self.assertEqual(c.number_format, 'DD/MM/YYYY')

    def test_link_do_dou_clicavel(self):
        c = self._celula(self._gerar()['Pautas'], 'Link DOU')
        self.assertEqual(c.value, 'Abrir no DOU')
        self.assertTrue(c.hyperlink.target.startswith('http://pesquisa.in.gov.br/'))

    def test_autofiltro_cobre_a_tabela_toda(self):
        wb = self._gerar()
        self.assertEqual(wb['Pautas'].auto_filter.ref, 'A1:N4')      # 3 processos + cabeçalho
        self.assertEqual(wb['Clientes na Pauta'].auto_filter.ref, 'A1:N2')
        # nada abaixo da tabela atrapalhando o filtro (a legenda foi para outra aba)
        self.assertEqual(wb['Pautas'].max_row, 4)



def _xml_parte(pasta, nome, texto_html, pagina):
    """XML no formato do INLabs com uma parte de pauta."""
    categoria = ('Ministério da Fazenda/Secretaria Especial da Receita Federal do Brasil/'
                 'Delegacia de Julgamento da Receita Federal do Brasil 09/Décima Quinta Turma')
    conteudo = (f'<xml><article artType="Pauta" pubDate="06/10/2026" artCategory="{categoria}" '
                f'pdfPage="http://pesquisa.in.gov.br/x?pagina={pagina}"><body>'
                f'<Texto><![CDATA[{texto_html}]]></Texto></body></article></xml>')
    caminho = os.path.join(pasta, nome)
    with open(caminho, 'w', encoding='utf-8') as fp:
        fp.write(conteudo)
    return caminho


def _proc(n, numero):
    return f'<p>{n} - Processo nº: {numero} - Recorrente: EMPRESA {n} LTDA e Interessado: FAZENDA NACIONAL</p>'


class TestPautaEmVariasPartes(unittest.TestCase):
    """Pauta grande dividida pelo INLabs em vários XMLs (caso real de 06/10/2026:
    DRJ09 15ª Turma em 17 partes; só as partes que abrem sessão têm cabeçalho)."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        d = self.tmp.name
        self.partes = [
            _xml_parte(d, '515_20261006_99-1.xml',
                       '<p>DIA 15 de Outubro de 2026 ÀS 09:00 HORAS</p><p>Relator(a): FULANA DE TAL</p>'
                       + _proc(1, '21020.539502/2022-01') + _proc(2, '40268.117758/2022-02'), 20),
            _xml_parte(d, '515_20261006_99-2.xml',
                       _proc(3, '91783.908470/2022-03') + _proc(4, '07661.771159/2022-04'), 22),
            # continuação da sessão anterior e, no meio, uma sessão nova
            _xml_parte(d, '515_20261006_99-10.xml',
                       _proc(5, '21249.985698/2022-05')
                       + '<p>DIA 15 de Outubro de 2026 ÀS 13:30 HORAS</p><p>Relator(a): BELTRANA</p>'
                       + _proc(1, '47896.118367/2022-06'), 35),
        ]
        self.solto = _xml_parte(d, '515_20261006_77.xml', _proc(1, '36576.615654/2022-07'), 40)

    def test_agrupa_partes_em_ordem_numerica(self):
        grupos = pauta.agrupar_partes(list(reversed(self.partes)) + [self.solto])
        self.assertIn([self.solto], grupos)
        self.assertIn(self.partes, grupos)   # -1, -2, -10 (não -1, -10, -2)
        self.assertEqual(len(grupos), 2)

    def test_partes_juntas_herdam_data_horario_e_relator(self):
        avisos = []
        with contextlib.redirect_stdout(io.StringIO()):
            linhas = pauta.parse_xml_pauta(self.partes, avisos)
        self.assertEqual(avisos, [])
        resumo = [(l['Número do Processo'][-7:], l['Data de Julgamento'], l['Horário'], l['Relator(a)'])
                  for l in linhas]
        self.assertEqual(resumo, [
            ('2022-01', '15 de Outubro de 2026', '09:00', 'FULANA DE TAL'),
            ('2022-02', '15 de Outubro de 2026', '09:00', 'FULANA DE TAL'),
            ('2022-03', '15 de Outubro de 2026', '09:00', 'FULANA DE TAL'),
            ('2022-04', '15 de Outubro de 2026', '09:00', 'FULANA DE TAL'),
            ('2022-05', '15 de Outubro de 2026', '09:00', 'FULANA DE TAL'),
            ('2022-06', '15 de Outubro de 2026', '13:30', 'BELTRANA'),
        ])

    def test_texto_antes_do_cabecalho_nao_e_descartado(self):
        # parte lida sozinha (sem as anteriores): o processo 5, antes do
        # cabeçalho, sai sem data, mas sai — antes era perdido
        avisos = []
        with contextlib.redirect_stdout(io.StringIO()):
            linhas = pauta.parse_xml_pauta(self.partes[2], avisos)
        self.assertEqual([l['Número do Processo'][-7:] for l in linhas], ['2022-05', '2022-06'])
        self.assertEqual(linhas[0]['Data de Julgamento'], '')
        self.assertEqual(avisos, [])


if __name__ == '__main__':
    unittest.main()
