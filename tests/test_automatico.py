# -*- coding: utf-8 -*-
"""
Testes do modo automático (agendamento, controle de dias, avisos), usando o
INLabs falso de test_download.py. Nada é enviado de verdade: e-mail,
notificação e launchctl são simulados.

Rodar (da pasta do projeto):
    python3 -m unittest discover tests -v
"""

import contextlib
import io
import json
import os
import plistlib
import shutil
import tempfile
import unittest
from datetime import date, datetime, timedelta
from unittest import mock

from openpyxl import Workbook

from test_download import BaseDownload, InlabsFalso, pauta

HOJE = date(2026, 9, 30)
PROCESSO_CLIENTE = '43106.903311/2006-41'   # está na pauta do ZIP falso (ALFA BRASIL NORDESTE)


def _estado(**kw):
    estado = {'inicio': HOJE.isoformat(), 'concluidos': [], 'avisados': {}}
    estado.update(kw)
    return estado


class TestDatasPendentes(unittest.TestCase):

    def test_primeira_execucao_so_hoje(self):
        with tempfile.TemporaryDirectory() as tmp:
            estado = pauta.carregar_estado(tmp, HOJE)
        self.assertEqual(pauta.datas_pendentes(estado, HOJE), ['2026-09-30'])

    def test_recupera_dias_nao_concluidos(self):
        estado = _estado(inicio='2026-09-25', concluidos=['2026-09-25', '2026-09-26', '2026-09-28'])
        self.assertEqual(pauta.datas_pendentes(estado, HOJE), ['2026-09-27', '2026-09-29', '2026-09-30'])

    def test_limite_de_dias_para_tras(self):
        estado = _estado(inicio='2026-01-01')
        pendentes = pauta.datas_pendentes(estado, HOJE)
        self.assertEqual(pendentes[0], (HOJE - timedelta(days=pauta.DIAS_RECUPERAR)).isoformat())
        self.assertEqual(pendentes[-1], HOJE.isoformat())

    def test_estado_esquece_historico_antigo(self):
        with tempfile.TemporaryDirectory() as tmp:
            estado = _estado(concluidos=['2026-01-01', '2026-09-29'],
                             avisados={'2026-01-01': ['x'], '2026-09-29': ['y']})
            pauta.salvar_estado(tmp, estado, HOJE)
            salvo = pauta.carregar_estado(tmp, HOJE)
        self.assertEqual(salvo['concluidos'], ['2026-09-29'])
        self.assertEqual(salvo['avisados'], {'2026-09-29': ['y']})


class TestRodarAutomatico(BaseDownload):

    def setUp(self):
        super().setUp()
        wb = Workbook()
        wb.active.append(['Nome', 'Processo'])
        wb.active.append(['Alfa Brasil Nordeste Ltda', PROCESSO_CLIENTE])
        wb.save(os.path.join(self.tmp.name, 'Clientes monitorados.xlsx'))
        # (patch a patch: mock.patch.stopall desligaria também o INLabs falso da classe)
        self.email = self.patch(pauta, 'enviar_email')
        self.notif = self.patch(pauta, 'notificar_mac')
        self.patch(pauta, 'ler_config_email',
                   return_value={'usuario': 'eu@b.com', 'destinatarios': ['eu@b.com', 'advogada@b.com']})

    def patch(self, alvo, nome, **kw):
        p = mock.patch.object(alvo, nome, **kw)
        self.addCleanup(p.stop)
        return p.start()

    def rodar(self, hoje=HOJE, hora=9, **comportamento):
        InlabsFalso.comportamento = comportamento
        saida = io.StringIO()
        agora = datetime(hoje.year, hoje.month, hoje.day, hora)
        with contextlib.redirect_stdout(saida), contextlib.redirect_stderr(io.StringIO()):
            codigo = pauta.rodar_automatico(self.tmp.name, hoje, agora)
        return codigo, saida.getvalue()

    def estado(self):
        with open(os.path.join(self.tmp.name, pauta.ARQ_ESTADO), encoding='utf-8') as fp:
            return json.load(fp)

    def test_cliente_na_pauta_avisa_por_email_e_notificacao(self):
        codigo, saida = self.rodar(DO1='ok')
        self.assertEqual(codigo, 0, saida)
        self.email.assert_called_once()
        _, assunto, corpo, anexos = self.email.call_args[0]
        self.assertEqual(assunto, 'Pautas CARF/DRJ: 1 cliente(s) na pauta')
        self.assertIn('Alfa Brasil Nordeste Ltda', corpo)
        self.assertIn(PROCESSO_CLIENTE, corpo)
        self.assertIn('Julgamento: 08/10/2026 09:00', corpo)
        self.assertIn('http://pesquisa.in.gov.br/', corpo)
        self.assertEqual([os.path.basename(a) for a in anexos], ['pautas_2026-09-30.xlsx'])
        self.notif.assert_called_once()
        self.assertIn('Alfa Brasil Nordeste Ltda', self.notif.call_args[0][1])
        self.assertEqual(self.estado()['avisados'], {'2026-09-30': [PROCESSO_CLIENTE]})
        html = self.email.call_args.kwargs['html']
        self.assertIn('Alfa Brasil Nordeste Ltda', html)
        self.assertIn('href="http://pesquisa.in.gov.br/', html)
        self.assertIn('08/10/2026', html)

    def test_lista_de_clientes_do_google_sheets(self):
        origem = os.path.join(self.tmp.name, 'Clientes monitorados.xlsx')
        os.rename(origem, os.path.join(self.tmp.name, 'lista_google.xlsx'))   # não está mais na pasta

        def baixar(planilha_id, destino):
            self.assertEqual(planilha_id, 'ID123')
            shutil.copy(os.path.join(self.tmp.name, 'lista_google.xlsx'), destino)
            return destino
        self.patch(pauta, 'baixar_clientes_google', side_effect=baixar)
        with mock.patch.dict(os.environ, {'CLIENTES_PLANILHA_ID': 'ID123'}):
            codigo, saida = self.rodar(DO1='ok')
        self.assertEqual(codigo, 0, saida)
        self.assertIn('Alfa Brasil Nordeste Ltda', self.email.call_args[0][2])

    def test_falha_no_google_sheets_e_problema(self):
        self.patch(pauta, 'baixar_clientes_google', side_effect=ValueError('sem acesso'))
        with mock.patch.dict(os.environ, {'CLIENTES_PLANILHA_ID': 'ID123'}):
            codigo, _ = self.rodar(DO1='ok')
        self.assertEqual(codigo, 2)
        self.assertIn('Google Sheets (sem acesso)', self.email.call_args[0][2])

    def test_link_da_execucao_no_github(self):
        env = {'GITHUB_RUN_ID': '42', 'GITHUB_REPOSITORY': 'dono/repo', 'GITHUB_SERVER_URL': 'https://github.com'}
        with mock.patch.dict(os.environ, env):
            self.rodar(DO1='ok_sem_pauta')
        self.assertIn('https://github.com/dono/repo/actions/runs/42', self.email.call_args[0][2])

    def test_html_escapa_nomes(self):
        html = pauta._html_aviso('T', [], [('2026-09-30', 'falha <b>&</b>')], [], '/pasta', False)
        self.assertIn('falha &lt;b&gt;&amp;&lt;/b&gt;', html)

    def test_nao_avisa_de_novo_o_mesmo_cliente_no_mesmo_dia(self):
        self.rodar(DO1='ok')                       # 8h
        self.email.reset_mock(); self.notif.reset_mock()
        codigo, saida = self.rodar(DO1='ok')       # 18h
        self.assertEqual(codigo, 0)
        self.email.assert_not_called()
        self.notif.assert_not_called()
        self.assertIn('Nada novo desde o último aviso', saida)

    def test_dia_sem_cliente_manda_resumo_com_a_planilha(self):
        wb = Workbook()
        wb.active.append(['Nome', 'Processo'])
        wb.active.append(['Outra Empresa Ltda', '99999.999999/9999-99'])
        wb.save(os.path.join(self.tmp.name, 'Clientes monitorados.xlsx'))
        codigo, _ = self.rodar(DO1='ok')
        self.assertEqual(codigo, 0)
        self.email.assert_called_once()
        _, assunto, corpo, anexos = self.email.call_args[0]
        self.assertEqual(assunto, 'Pautas CARF/DRJ de 30/09/2026: nenhum cliente na pauta')
        self.assertIn('Nenhum cliente monitorado', corpo)
        self.assertIn('PAUTAS NOVAS', corpo)
        self.assertEqual([os.path.basename(a) for a in anexos], ['pautas_2026-09-30.xlsx'])
        self.notif.assert_called_once()

    def test_dia_sem_pautas_manda_resumo_sem_anexo(self):
        self.rodar(DO1='ok_sem_pauta')
        _, assunto, corpo, anexos = self.email.call_args[0]
        self.assertIn('nenhum cliente na pauta', assunto)
        self.assertIn('Não houve pautas', corpo)
        self.assertEqual(anexos, [])

    def test_depois_das_20h_manda_so_para_mim_e_nao_salva_o_estado(self):
        self.rodar(hora=22, DO1='ok')                               # 18h atrasada pelo GitHub
        self.email.assert_called_once()
        self.assertEqual(self.email.call_args[0][0]['destinatarios'], ['eu@b.com'])
        self.assertFalse(os.path.exists(os.path.join(self.tmp.name, pauta.ARQ_ESTADO)))
        self.email.reset_mock()
        self.rodar(hoje=HOJE + timedelta(days=1), DO1='ok')         # 8h30 do dia seguinte: todo mundo
        self.assertEqual(self.email.call_args[0][0]['destinatarios'], ['eu@b.com', 'advogada@b.com'])
        self.assertIn('Alfa Brasil Nordeste Ltda', self.email.call_args[0][2])

    def test_sem_novidade_so_o_primeiro_aviso_do_dia(self):
        self.rodar(DO1='ok_sem_pauta')                              # 8h30
        self.email.reset_mock(); self.notif.reset_mock()
        _, saida = self.rodar(DO1='ok_sem_pauta')                   # 13h
        self.email.assert_not_called()
        self.notif.assert_not_called()
        self.assertIn('Nada novo desde o último aviso', saida)
        self.rodar(hoje=HOJE + timedelta(days=1), DO1='ok_sem_pauta')   # 8h30 do dia seguinte
        self.email.assert_called_once()

    def test_mesmas_pautas_a_tarde_nao_avisam_de_novo(self):
        self.rodar(DO1='ok')                                        # 8h30: cliente avisado
        self.email.reset_mock()
        self.rodar(DO1='ok')                                        # 13h: nada novo
        self.email.assert_not_called()

    def test_pauta_nova_a_tarde_avisa_mesmo_sem_cliente(self):
        wb = Workbook()
        wb.active.append(['Nome', 'Processo'])
        wb.active.append(['Outra Empresa Ltda', '99999.999999/9999-99'])
        wb.save(os.path.join(self.tmp.name, 'Clientes monitorados.xlsx'))
        self.rodar(DO1='ok_sem_pauta')                              # 8h30: sem pautas
        self.email.reset_mock()
        self.rodar(DO1='ok')                                        # 13h: saiu pauta
        _, assunto, corpo, anexos = self.email.call_args[0]
        self.assertRegex(assunto, r'^Pautas CARF/DRJ de 30/09/2026: \d+ pauta\(s\) nova\(s\), nenhum cliente$')
        self.assertIn('DOU de 30/09/2026:', corpo)
        self.assertEqual([os.path.basename(a) for a in anexos], ['pautas_2026-09-30.xlsx'])

    def test_pauta_nova_a_tarde_com_cliente(self):
        self.rodar(DO1='ok_sem_pauta')                              # 8h30
        self.email.reset_mock()
        self.rodar(DO1='ok')                                        # 18h
        _, assunto, corpo, _ = self.email.call_args[0]
        self.assertEqual(assunto, 'Pautas CARF/DRJ: 1 cliente(s) na pauta')
        self.assertIn('Alfa Brasil Nordeste Ltda', corpo)

    def test_hoje_nao_e_concluido_e_ontem_e_concluido_no_dia_seguinte(self):
        self.rodar(DO1='ok_sem_pauta')
        self.assertEqual(self.estado()['concluidos'], [])          # hoje ainda pode ter extra
        self.rodar(hoje=HOJE + timedelta(days=1), DO1='ok_sem_pauta')
        self.assertEqual(self.estado()['concluidos'], ['2026-09-30'])

    def test_recupera_dia_em_que_o_mac_ficou_desligado(self):
        self.rodar(DO1='ok_sem_pauta')                             # 30/09
        _, saida = self.rodar(hoje=HOJE + timedelta(days=3), DO1='ok_sem_pauta')   # só em 03/10
        self.assertIn('Dias a processar: 30/09/2026, 01/10/2026, 02/10/2026, 03/10/2026', saida)
        self.assertEqual(self.estado()['concluidos'], ['2026-09-30', '2026-10-01', '2026-10-02'])

    def test_erro_no_download_avisa_e_nao_conclui_o_dia(self):
        self.rodar(DO1='ok_sem_pauta')
        codigo, _ = self.rodar(hoje=HOJE + timedelta(days=1), DO1='ok_sem_pauta', DO2='500')
        self.assertEqual(codigo, 2)
        _, assunto, corpo, _ = self.email.call_args[0]
        self.assertEqual(assunto, 'Pautas CARF/DRJ: 2 problema(s)')   # DO2 falhou nos dois dias
        self.assertIn('DOU de 30/09/2026: DO2 não foi baixado', corpo)
        self.assertEqual(self.estado()['concluidos'], [])

    def test_aviso_so_de_erro_vai_so_para_mim(self):
        self.rodar(DO1='ok_sem_pauta', DO2='500')                   # 8h30: só erro
        self.assertEqual(self.email.call_args[0][0]['destinatarios'], ['eu@b.com'])
        self.assertNotIn('resumo_enviado', self.estado())
        self.email.reset_mock()
        self.rodar(DO1='ok_sem_pauta')                              # 13h: deu certo, resumo para todos
        self.assertEqual(self.email.call_args[0][0]['destinatarios'], ['eu@b.com', 'advogada@b.com'])

    def test_erro_junto_com_cliente_vai_para_todos(self):
        self.rodar(DO1='ok', DO2='500')
        self.assertEqual(self.email.call_args[0][0]['destinatarios'], ['eu@b.com', 'advogada@b.com'])
        self.assertIn('DO2 não foi baixado', self.email.call_args[0][2])

    def test_mesmo_problema_so_um_aviso_por_dia(self):
        codigo, _ = self.rodar(DO1='ok_sem_pauta', DO2='500')       # 8h07
        self.assertEqual(codigo, 2)
        self.assertIn('problema', self.email.call_args[0][1])
        self.email.reset_mock()
        codigo, saida = self.rodar(DO1='ok_sem_pauta', DO2='500')   # 8h22: mesmo problema
        self.assertEqual(codigo, 2)
        self.email.assert_not_called()
        self.assertIn('Nada novo desde o último aviso', saida)
        self.rodar(hoje=HOJE + timedelta(days=1), DO1='ok_sem_pauta', DO2='500')   # dia seguinte avisa
        self.assertIn('problema', self.email.call_args[0][1])

    def test_dia_sem_dou_nao_e_problema(self):
        codigo, _ = self.rodar()   # tudo 404 (feriado)
        self.assertEqual(codigo, 0)
        _, assunto, corpo, anexos = self.email.call_args[0]
        self.assertIn('nenhum cliente na pauta', assunto)   # só o resumo, sem "problema"
        self.assertIn('Não houve pautas', corpo)

    def test_fim_de_semana_nao_e_problema_e_e_concluido(self):
        # sexta 02/10 → segunda 05/10: sábado e domingo vêm como página de listagem
        sexta = date(2026, 10, 2)
        self.rodar(hoje=sexta, DO1='ok_sem_pauta')
        self.email.reset_mock()
        codigo, saida = self.rodar(hoje=date(2026, 10, 5), DO1='listagem', DO2='listagem', DO3='listagem')
        self.assertEqual(codigo, 0, saida)
        self.assertNotIn('problema', self.email.call_args[0][1])
        self.assertEqual(self.estado()['concluidos'], ['2026-10-02', '2026-10-03', '2026-10-04'])

    def test_senha_errada_avisa(self):
        with mock.patch.dict(os.environ, {'INLABS_SENHA': 'errada'}):
            codigo, _ = self.rodar(DO1='ok')
        self.assertEqual(codigo, 2)
        self.assertIn('recusou o login', self.email.call_args[0][2])

    def test_falha_no_email_aparece_na_notificacao(self):
        self.email.side_effect = OSError('sem internet')
        self.rodar(DO1='ok')
        self.assertIn('E-mail NÃO enviado: sem internet', self.notif.call_args[0][1])

    def test_erro_inesperado_nao_derruba_os_outros_dias(self):
        self.rodar(DO1='ok_sem_pauta')
        original = pauta.executar
        def executar_quebrando_no_primeiro(args, pasta):
            if args.baixar == '2026-09-30':
                raise RuntimeError('disco cheio')
            return original(args, pasta)
        with mock.patch.object(pauta, 'executar', side_effect=executar_quebrando_no_primeiro):
            codigo, saida = self.rodar(hoje=HOJE + timedelta(days=1), DO1='ok')
        self.assertEqual(codigo, 2)
        self.assertIn('----- 01/10/2026 -----', saida)
        _, assunto, corpo, _ = self.email.call_args[0]
        self.assertEqual(assunto, 'Pautas CARF/DRJ: 1 cliente(s) na pauta e 1 problema(s)')
        self.assertIn('erro inesperado: RuntimeError: disco cheio', corpo)


class TestConfigEmail(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.caminho = os.path.join(self.tmp.name, pauta.ARQ_EMAIL)

    def test_cria_modelo_na_primeira_vez(self):
        with self.assertRaisesRegex(ValueError, 'criei'):
            pauta.ler_config_email(self.tmp.name)
        self.assertEqual(os.stat(self.caminho).st_mode & 0o777, 0o600)

    def test_modelo_nao_preenchido(self):
        with open(self.caminho, 'w') as fp:
            fp.write(pauta.MODELO_EMAIL)
        with self.assertRaisesRegex(ValueError, 'preencha'):
            pauta.ler_config_email(self.tmp.name)

    def test_le_varios_destinatarios(self):
        with open(self.caminho, 'w') as fp:
            fp.write('[email]\nservidor = smtp.gmail.com\nporta = 465\nusuario = eu@x.com\n'
                     'senha = abc%def\ndestinatarios = a@x.com, b@x.com\n')
        config = pauta.ler_config_email(self.tmp.name)
        self.assertEqual(config['destinatarios'], ['a@x.com', 'b@x.com'])
        self.assertEqual(config['senha'], 'abc%def')


class TestEnviarEmail(unittest.TestCase):

    CONFIG = {'servidor': 'smtp.gmail.com', 'porta': 465, 'usuario': 'eu@x.com',
              'senha': 's', 'destinatarios': ['a@x.com', 'b@x.com']}

    def test_envia_com_anexo(self):
        with tempfile.TemporaryDirectory() as tmp:
            anexo = os.path.join(tmp, 'pautas_2026-09-30.xlsx')
            with open(anexo, 'wb') as fp:
                fp.write(b'conteudo')
            with mock.patch('smtplib.SMTP_SSL') as smtp:
                pauta.enviar_email(self.CONFIG, 'Assunto', 'Corpo', [anexo])
        conexao = smtp.return_value
        conexao.login.assert_called_once_with('eu@x.com', 's')
        msg = conexao.send_message.call_args[0][0]
        self.assertEqual(msg['To'], 'a@x.com, b@x.com')
        self.assertEqual([a.get_filename() for a in msg.iter_attachments()], ['pautas_2026-09-30.xlsx'])

    def test_config_pelas_variaveis_de_ambiente(self):
        env = {'EMAIL_USUARIO': 'eu@x.com', 'EMAIL_SENHA': 's', 'EMAIL_DESTINATARIOS': 'a@x.com, b@x.com'}
        with tempfile.TemporaryDirectory() as tmp, mock.patch.dict(os.environ, env):
            config = pauta.ler_config_email(tmp)
            self.assertFalse(os.path.exists(os.path.join(tmp, pauta.ARQ_EMAIL)))
        self.assertEqual(config, {'servidor': 'smtp.gmail.com', 'porta': 465, 'usuario': 'eu@x.com',
                                  'senha': 's', 'destinatarios': ['a@x.com', 'b@x.com']})

    def test_envia_html_texto_e_anexo(self):
        with tempfile.TemporaryDirectory() as tmp:
            anexo = os.path.join(tmp, 'pautas_2026-09-30.xlsx')
            with open(anexo, 'wb') as fp:
                fp.write(b'conteudo')
            with mock.patch('smtplib.SMTP_SSL') as smtp:
                pauta.enviar_email(self.CONFIG, 'Assunto', 'Corpo', [anexo], html='<p>Corpo</p>')
        msg = smtp.return_value.send_message.call_args[0][0]
        self.assertEqual(msg.get_body(('plain',)).get_content().strip(), 'Corpo')
        self.assertIn('<p>Corpo</p>', msg.get_body(('html',)).get_content())
        self.assertEqual([a.get_filename() for a in msg.iter_attachments()], ['pautas_2026-09-30.xlsx'])

    def test_senha_normal_em_vez_de_senha_de_app_explica_o_erro(self):
        import smtplib
        with mock.patch('smtplib.SMTP_SSL') as smtp:
            smtp.return_value.login.side_effect = smtplib.SMTPServerDisconnected('Connection unexpectedly closed')
            with self.assertRaisesRegex(ValueError, 'senha de app'):
                pauta.enviar_email(self.CONFIG, 'A', 'C')

    def test_porta_587_usa_starttls(self):
        with mock.patch('smtplib.SMTP') as smtp:
            pauta.enviar_email(dict(self.CONFIG, porta=587), 'A', 'C')
        smtp.return_value.starttls.assert_called_once()


class TestNotificacaoEAgendamento(unittest.TestCase):

    def test_notificacao_passa_texto_como_argumento(self):
        with mock.patch('sys.platform', 'darwin'), mock.patch('subprocess.run') as run, \
             contextlib.redirect_stdout(io.StringIO()):
            pauta.notificar_mac('Título', 'EMPRESA "ASPAS" & AÇÚCAR')
        cmd = run.call_args[0][0]
        self.assertEqual(cmd[0], 'osascript')
        self.assertEqual(cmd[-2:], ['Título', 'EMPRESA "ASPAS" & AÇÚCAR'])
        self.assertNotIn('ASPAS', ' '.join(cmd[:-2]))

    def test_instala_plist_de_segunda_a_sexta_8h30_13h_18h(self):
        with tempfile.TemporaryDirectory() as tmp:
            plist = os.path.join(tmp, 'LaunchAgents', 'x.plist')
            with mock.patch.object(pauta, '_caminho_plist', return_value=plist), \
                 mock.patch('subprocess.run', return_value=mock.Mock(returncode=0)) as run, \
                 contextlib.redirect_stdout(io.StringIO()):
                codigo = pauta.instalar_agendamento(tmp)
            with open(plist, 'rb') as fp:
                config = plistlib.load(fp)
        self.assertEqual(codigo, 0)
        self.assertEqual(config['ProgramArguments'][-1], '--automatico')
        self.assertEqual(len(config['StartCalendarInterval']), 15)
        self.assertEqual({(i['Weekday'], i['Hour'], i['Minute']) for i in config['StartCalendarInterval']},
                         {(d, h, m) for d in range(1, 6) for h, m in ((8, 30), (13, 0), (18, 0))})
        self.assertTrue(config['StandardOutPath'].endswith(os.path.join('logs', 'automatico.log')))
        self.assertIn('bootstrap', run.call_args[0][0])


if __name__ == '__main__':
    unittest.main()
