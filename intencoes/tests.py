import os
from datetime import date, time, timedelta
from django.test import TestCase, Client
from django.urls import reverse
from django.utils import timezone
from .models import Missa, IntencaoMissa, Pagamento, RegistroCaixa
from .docx_generator import gerar_docx_missa
from .docx_importer import importar_docx_para_missa
from .excel_generator import gerar_excel_missa
from .pdf_generator import gerar_pdf_missa
from docx import Document as DocxDocument
from openpyxl import load_workbook


class IntencoesMissaTests(TestCase):
    def setUp(self):
        self.client = Client()
        agora = timezone.localtime()
        data_futura = (agora + timedelta(days=2)).date()
        self.missa = Missa.objects.create(
            data=data_futura,
            horario=time(19, 0)
        )

    def test_painel_missas_view(self):
        response = self.client.get(reverse('painel_missas'))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Painel de Celebrações')
        self.assertContains(response, self.missa.data.strftime('%d/%m/%Y'))

    def test_detalhes_missa_view(self):
        response = self.client.get(reverse('detalhes_missa', kwargs={'missa_id': self.missa.id}))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Intenções por Categoria Litúrgica')
        self.assertContains(response, 'Controle do Caixa da Sessão')
        self.assertContains(response, 'Doações do Troco')

    def test_criar_missa_view(self):
        agora = timezone.localtime()
        data_nova = (agora + timedelta(days=5)).date()
        response = self.client.post(reverse('criar_missa'), {
            'data': data_nova.strftime('%Y-%m-%d'),
            'horario': '10:00'
        })
        self.assertEqual(response.status_code, 302)
        self.assertTrue(Missa.objects.filter(horario=time(10, 0), data=data_nova).exists())

    def test_timer_colorido(self):
        agora = timezone.localtime()

        # 1. Verde (> 1 dia)
        m_verde = Missa.objects.create(data=(agora + timedelta(days=3)).date(), horario=time(18, 0))
        self.assertEqual(m_verde.status_timer['cor'], 'verde')

        # 2. Amarelo (< 24h e > 10 min)
        dt_amarelo = agora + timedelta(hours=5)
        m_amarelo = Missa.objects.create(data=dt_amarelo.date(), horario=time(dt_amarelo.hour, dt_amarelo.minute))
        self.assertEqual(m_amarelo.status_timer['cor'], 'amarelo')

        # 3. Vermelho (<= 10 min)
        dt_vermelho = agora + timedelta(minutes=5)
        m_vermelho = Missa.objects.create(data=dt_vermelho.date(), horario=time(dt_vermelho.hour, dt_vermelho.minute))
        self.assertEqual(m_vermelho.status_timer['cor'], 'vermelho')

        # 4. Cinza (+ 3h encerrada)
        dt_cinza = agora - timedelta(hours=4)
        m_cinza = Missa.objects.create(data=dt_cinza.date(), horario=time(dt_cinza.hour, dt_cinza.minute))
        self.assertEqual(m_cinza.status_timer['cor'], 'cinza')
        self.assertTrue(m_cinza.esta_bloqueada)

    def test_bloqueio_apos_3_horas(self):
        agora = timezone.localtime()
        dt_antiga = agora - timedelta(hours=4)
        missa_bloqueada = Missa.objects.create(
            data=dt_antiga.date(),
            horario=time(dt_antiga.hour, dt_antiga.minute)
        )
        self.assertTrue(missa_bloqueada.esta_bloqueada)

        response = self.client.post(
            reverse('cadastrar_intencao', kwargs={'missa_id': missa_bloqueada.id}),
            {'nome_intencao': 'teste', 'tipo': 'falecido'}
        )
        self.assertEqual(response.status_code, 302)
        self.assertRedirects(response, reverse('detalhes_missa', kwargs={'missa_id': missa_bloqueada.id}))
        self.assertEqual(missa_bloqueada.intencoes.count(), 0)

        intencao = IntencaoMissa.objects.create(missa=missa_bloqueada, nome_intencao='Original', tipo='falecido')
        response_edit = self.client.post(
            reverse('editar_intencao', kwargs={'intencao_id': intencao.id}),
            {'nome_intencao': 'Modificado', 'tipo': 'saude'}
        )
        self.assertEqual(response_edit.status_code, 302)
        intencao.refresh_from_db()
        self.assertEqual(intencao.nome_intencao, 'Original')

        response_del = self.client.post(reverse('excluir_intencao', kwargs={'intencao_id': intencao.id}))
        self.assertEqual(response_del.status_code, 302)
        self.assertEqual(missa_bloqueada.intencoes.count(), 1)

        res_docx = self.client.get(reverse('baixar_docx', kwargs={'missa_id': missa_bloqueada.id}))
        self.assertEqual(res_docx.status_code, 200)
        res_caixa = self.client.get(reverse('caixa_missa', kwargs={'missa_id': missa_bloqueada.id}))
        self.assertEqual(res_caixa.status_code, 200)

    def test_lote_pendente_e_auditoria(self):
        # 1. Cadastra intenções na rodada
        self.client.post(
            reverse('cadastrar_intencao', kwargs={'missa_id': self.missa.id}),
            {'nome_intencao': 'joao silva', 'tipo': 'falecido'}
        )
        self.client.post(
            reverse('cadastrar_intencao', kwargs={'missa_id': self.missa.id}),
            {'nome_intencao': 'maria sousa', 'tipo': 'saude'}
        )
        self.assertEqual(self.missa.intencoes.count(), 2)

        # 2. Suspende e salva como pendente
        response_salvar = self.client.get(reverse('salvar_lote_pendente', kwargs={'missa_id': self.missa.id}))
        self.assertEqual(response_salvar.status_code, 302)

        # Verifica auditoria gerada
        reg_pendente = RegistroCaixa.objects.filter(missa=self.missa, texto__contains='PAGAMENTO PENDENTE').first()
        self.assertIsNotNone(reg_pendente)
        self.assertIn('Joao Silva', reg_pendente.texto)

        # Verifica que NÃO entra no Word (DOCX)
        buffer = gerar_docx_missa(self.missa)
        # Como as 2 intenções são pendentes, o docx não deve conter seus nomes
        self.assertEqual(self.missa.intencoes_sessao_count, 0)
        self.assertEqual(self.missa.total_sessao, 0)

        # 3. Quita o lote a partir do menu da missa
        lote_pendente = self.missa.lotes_pendentes.first()
        self.assertIsNotNone(lote_pendente)

        response_pagar = self.client.post(
            reverse('pagamento_lote', kwargs={'missa_id': self.missa.id, 'pagamento_id': lote_pendente.id}),
            {
                'forma': 'pix',
                'valor_recebido': '6.00',
                'doacao': '0.00',
                'nome_pix': 'carlos alberto da costa'  # testa .title()
            }
        )
        self.assertEqual(response_pagar.status_code, 302)

        lote_pendente.refresh_from_db()
        self.assertEqual(lote_pendente.status, 'pago')
        self.assertEqual(lote_pendente.nome_pix, 'Carlos Alberto Da Costa')

        # Verifica auditoria de pagamento confirmado
        reg_confirmado = RegistroCaixa.objects.filter(missa=self.missa, texto__contains='CONFIRMADO').first()
        self.assertIsNotNone(reg_confirmado)
        self.assertIn('Carlos Alberto Da Costa', reg_confirmado.texto)

        # Agora as intenções entram no total da sessão e no Word
        self.assertEqual(self.missa.intencoes_sessao_count, 2)
        self.assertEqual(self.missa.total_sessao, 6.00)

    def test_cadastrar_intencao_e_remover_pendente(self):
        response = self.client.post(
            reverse('cadastrar_intencao', kwargs={'missa_id': self.missa.id}),
            {'nome_intencao': 'joao silva', 'tipo': 'falecido'}
        )
        self.assertEqual(response.status_code, 302)
        intencao = IntencaoMissa.objects.first()
        self.assertIsNotNone(intencao)
        self.assertEqual(intencao.nome_intencao, 'Joao Silva')

        response = self.client.get(
            reverse('remover_intencao_pendente', kwargs={'missa_id': self.missa.id, 'intencao_id': intencao.id})
        )
        self.assertEqual(response.status_code, 302)
        self.assertEqual(IntencaoMissa.objects.count(), 0)

    def test_cancelar_rodada(self):
        self.client.post(
            reverse('cadastrar_intencao', kwargs={'missa_id': self.missa.id}),
            {'nome_intencao': 'maria sousa', 'tipo': 'saude'}
        )
        self.assertEqual(IntencaoMissa.objects.count(), 1)

        response = self.client.get(reverse('cancelar_rodada', kwargs={'missa_id': self.missa.id}))
        self.assertEqual(response.status_code, 302)
        self.assertEqual(IntencaoMissa.objects.count(), 0)

    def test_validacao_impede_pagamento_sem_forma(self):
        self.client.post(
            reverse('cadastrar_intencao', kwargs={'missa_id': self.missa.id}),
            {'nome_intencao': 'antonio carlos', 'tipo': 'falecido'}
        )
        response = self.client.post(
            reverse('pagamento', kwargs={'missa_id': self.missa.id}),
            {
                'forma': '',
                'valor_recebido': '3.00',
                'doacao': '0.00',
                'nome_pix': ''
            }
        )
        self.assertEqual(response.status_code, 200)
        self.assertFormError(response.context['form'], 'forma', 'É obrigatório selecionar uma forma de pagamento (Dinheiro ou Pix).')

    def test_validacao_impede_pix_sem_nome_responsavel(self):
        self.client.post(
            reverse('cadastrar_intencao', kwargs={'missa_id': self.missa.id}),
            {'nome_intencao': 'antonio carlos', 'tipo': 'falecido'}
        )
        response = self.client.post(
            reverse('pagamento', kwargs={'missa_id': self.missa.id}),
            {
                'forma': 'pix',
                'valor_recebido': '3.00',
                'doacao': '0.00',
                'nome_pix': '   '
            }
        )
        self.assertEqual(response.status_code, 200)
        self.assertFormError(response.context['form'], 'nome_pix', 'É obrigatório informar o nome do responsável pelo Pix para conferência.')

    def test_fluxo_pagamento_dinheiro_com_troco_e_doacao(self):
        self.client.post(
            reverse('cadastrar_intencao', kwargs={'missa_id': self.missa.id}),
            {'nome_intencao': 'intencao 1', 'tipo': 'acao_gracas'}
        )
        self.client.post(
            reverse('cadastrar_intencao', kwargs={'missa_id': self.missa.id}),
            {'nome_intencao': 'intencao 2', 'tipo': 'honra'}
        )

        response = self.client.post(
            reverse('pagamento', kwargs={'missa_id': self.missa.id}),
            {
                'forma': 'dinheiro',
                'valor_recebido': '10.00',
                'doacao': '2.00',
                'nome_pix': ''
            }
        )
        self.assertEqual(response.status_code, 302)

        pagamento = Pagamento.objects.filter(status='pago').first()
        self.assertIsNotNone(pagamento)
        self.assertEqual(pagamento.valor_total, 6.00)
        self.assertEqual(pagamento.valor_recebido, 10.00)
        self.assertEqual(pagamento.doacao, 2.00)
        self.assertEqual(pagamento.troco, 2.00)

    def test_separacao_total_sessao_e_secretaria(self):
        p_sessao = Pagamento.objects.create(
            missa=self.missa, forma='dinheiro', status='pago', valor_total=6.00, valor_recebido=6.00
        )
        IntencaoMissa.objects.create(missa=self.missa, nome_intencao='Intencao 1', tipo='falecido', pagamento=p_sessao)
        IntencaoMissa.objects.create(missa=self.missa, nome_intencao='Intencao 2', tipo='saude', pagamento=p_sessao)

        p_sec = Pagamento.objects.create(
            missa=self.missa, forma='secretaria', status='pago', valor_total=9.00, valor_recebido=9.00
        )
        IntencaoMissa.objects.create(missa=self.missa, nome_intencao='Sec 1', tipo='niver', pagamento=p_sec)
        IntencaoMissa.objects.create(missa=self.missa, nome_intencao='Sec 2', tipo='honra', pagamento=p_sec)
        IntencaoMissa.objects.create(missa=self.missa, nome_intencao='Sec 3', tipo='acao_gracas', pagamento=p_sec)

        self.assertEqual(self.missa.intencoes_sessao_count, 2)
        self.assertEqual(self.missa.intencoes_secretaria_count, 3)
        self.assertEqual(self.missa.total_sessao, 6.00)
        self.assertEqual(self.missa.total_secretaria, 9.00)
        self.assertEqual(self.missa.total_arrecadado, 15.00)

    def test_importacao_docx_com_deteccao_automatica_e_ja_pago(self):
        caminho_docx = 'exemplo/intencoes.docx'
        if not os.path.exists(caminho_docx):
            caminho_docx = 'input/intencoes.docx'
        if os.path.exists(caminho_docx):
            missa_resultado, qtd, criada = importar_docx_para_missa(caminho_docx)
            self.assertEqual(missa_resultado.data, date(2026, 11, 7))
            self.assertEqual(missa_resultado.horario, time(19, 0))
            self.assertTrue(qtd > 0)

            intencoes = missa_resultado.intencoes.all()
            self.assertEqual(intencoes.count(), qtd)
            for item in intencoes:
                self.assertIsNotNone(item.pagamento)
                self.assertEqual(item.pagamento.forma, 'secretaria')
                self.assertEqual(item.pagamento.status, 'pago')

    def test_editar_e_excluir_intencao(self):
        intencao = IntencaoMissa.objects.create(
            missa=self.missa,
            nome_intencao='Nome Errado',
            tipo='falecido'
        )

        response = self.client.post(
            reverse('editar_intencao', kwargs={'intencao_id': intencao.id}),
            {'nome_intencao': 'Nome Corrigido', 'tipo': 'saude'}
        )
        self.assertEqual(response.status_code, 302)
        intencao.refresh_from_db()
        self.assertEqual(intencao.nome_intencao, 'Nome Corrigido')
        self.assertEqual(intencao.tipo, 'saude')

        response = self.client.post(reverse('excluir_intencao', kwargs={'intencao_id': intencao.id}))
        self.assertEqual(response.status_code, 302)
        self.assertEqual(IntencaoMissa.objects.filter(id=intencao.id).count(), 0)

    def test_gerar_docx_missa(self):
        p_pago = Pagamento.objects.create(missa=self.missa, forma='dinheiro', status='pago', valor_total=3.00, valor_recebido=3.00)
        IntencaoMissa.objects.create(
            missa=self.missa,
            nome_intencao='João Paulo II',
            tipo='honra',
            pagamento=p_pago
        )
        buffer = gerar_docx_missa(self.missa)
        self.assertTrue(buffer.getbuffer().nbytes > 1000)

        response = self.client.get(reverse('baixar_docx', kwargs={'missa_id': self.missa.id}))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response['Content-Type'],
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    def test_caixa_missa_view(self):
        Pagamento.objects.create(
            missa=self.missa,
            forma='dinheiro',
            status='pago',
            valor_total=6.00,
            valor_recebido=10.00,
            doacao=1.00
        )
        Pagamento.objects.create(
            missa=self.missa,
            forma='pix',
            status='pago',
            valor_total=3.00,
            valor_recebido=3.00,
            nome_pix='Carlos Eduardo'
        )

        response = self.client.get(reverse('caixa_missa', kwargs={'missa_id': self.missa.id}))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Carlos Eduardo')

        response_txt = self.client.get(reverse('caixa_missa', kwargs={'missa_id': self.missa.id}) + '?exportar=txt')
        self.assertEqual(response_txt.status_code, 200)
        txt = response_txt.content.decode('utf-8')
        self.assertIn('=== REGISTRO DE CAIXA - CATEDRAL DIVINO ESPÍRITO SANTO ===', txt)
        self.assertIn('Total de Intenções Pagas:', txt)
        self.assertIn('- Numero de intenções marcadas com dinheiro:', txt)
        self.assertIn('- Numero de inteções marcadas com pix:', txt)
        self.assertIn('- Numero de doações feitas:', txt)
        self.assertIn('- Numero de intenções marcadas na secretária:', txt)
        self.assertIn('Valor Total Arrecadado (Confirmado):', txt)
        self.assertIn('- Dinheiro em Caixa (Sessão):', txt)
        self.assertIn('- Pix Confirmado (Sessão):', txt)
        self.assertIn('- Doações Recebidas:', txt)
        self.assertIn('- Secretaria / Importado:', txt)
        self.assertIn('--- DETALHAMENTO DE PAGAMENTOS PIX CONFIRMADOS ---', txt)
        self.assertIn('Carlos Eduardo', txt)
        self.assertIn('--- TRILHA DE AUDITORIA E MOVIMENTAÇÃO (HISTÓRICO) ---', txt)

    def test_gerar_excel_missa(self):
        p_pago = Pagamento.objects.create(missa=self.missa, forma='dinheiro', status='pago', valor_total=3.00, valor_recebido=3.00)
        IntencaoMissa.objects.create(missa=self.missa, nome_intencao='Santa Teresinha', tipo='honra', pagamento=p_pago)

        buffer = gerar_excel_missa(self.missa)
        wb = load_workbook(buffer)
        self.assertIn('Intenções da Missa', wb.sheetnames)
        self.assertIn('Fechamento de Caixa', wb.sheetnames)

        response = self.client.get(reverse('baixar_excel', kwargs={'missa_id': self.missa.id}))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response['Content-Type'],
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    def test_gerar_pdf_missa(self):
        p_pago = Pagamento.objects.create(missa=self.missa, forma='dinheiro', status='pago', valor_total=3.00, valor_recebido=3.00)
        IntencaoMissa.objects.create(missa=self.missa, nome_intencao='São José', tipo='honra', pagamento=p_pago)

        buffer = gerar_pdf_missa(self.missa)
        conteudo_bytes = buffer.getvalue()
        self.assertTrue(len(conteudo_bytes) > 1000)
        self.assertTrue(conteudo_bytes.startswith(b'%PDF'))

        response = self.client.get(reverse('baixar_pdf', kwargs={'missa_id': self.missa.id}))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response['Content-Type'], 'application/pdf')
        self.assertTrue(response.get('Content-Disposition').endswith('.pdf"'))

    def test_asterisco_primeira_intencao_sessao_docx(self):
        # 1 intenção da secretaria
        p_sec = Pagamento.objects.create(missa=self.missa, forma='secretaria', status='pago', valor_total=3.00, valor_recebido=3.00)
        IntencaoMissa.objects.create(missa=self.missa, nome_intencao='Ana Secretaria', tipo='honra', pagamento=p_sec)

        # 2 intenções da sessão
        p_ses = Pagamento.objects.create(missa=self.missa, forma='dinheiro', status='pago', valor_total=6.00, valor_recebido=6.00)
        IntencaoMissa.objects.create(missa=self.missa, nome_intencao='Bruno Sessao Um', tipo='honra', pagamento=p_ses)
        IntencaoMissa.objects.create(missa=self.missa, nome_intencao='Claudio Sessao Dois', tipo='honra', pagamento=p_ses)

        buffer = gerar_docx_missa(self.missa)
        doc = DocxDocument(buffer)

        textos_tabela = []
        for t in doc.tables:
            for r in t.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        if p.text.strip():
                            textos_tabela.append(p.text.strip())

        # Primeira intenção da sessão deve ter '*'
        self.assertIn('Bruno Sessao Um *', textos_tabela)
        # Segunda intenção da sessão NÃO deve ter '*'
        self.assertIn('Claudio Sessao Dois', textos_tabela)
        self.assertNotIn('Claudio Sessao Dois *', textos_tabela)
        # Secretaria NÃO deve ter '*'
        self.assertIn('Ana Secretaria', textos_tabela)
        self.assertNotIn('Ana Secretaria *', textos_tabela)

    def test_importacao_docx_com_asterisco(self):
        import io
        from docx import Document

        doc = Document()
        doc.add_paragraph("Intenções de 15/11/2026 - Missa das 19h")
        t = doc.add_table(rows=5, cols=1)
        t.rows[0].cells[0].text = "HONRA E LOUVOR:"
        t.rows[1].cells[0].text = "Maria Aparecida Secretaria"
        t.rows[2].cells[0].text = "Jose Balcao Um *"
        t.rows[3].cells[0].text = "Antonio Balcao Dois"
        t.rows[4].cells[0].text = "Fim"

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)

        missa_res, qtd, criada = importar_docx_para_missa(buf)
        self.assertEqual(qtd, 3)

        # Secretaria
        self.assertEqual(missa_res.intencoes_secretaria_count, 1)
        it_sec = missa_res.intencoes.filter(nome_intencao__contains='Maria Aparecida').first()
        self.assertIsNotNone(it_sec)
        self.assertEqual(it_sec.pagamento.forma, 'secretaria')

        # Sessão (identificadas a partir do asterisco)
        self.assertEqual(missa_res.intencoes_sessao_count, 2)
        it_ses1 = missa_res.intencoes.filter(nome_intencao__contains='Jose Balcao Um').first()
        it_ses2 = missa_res.intencoes.filter(nome_intencao__contains='Antonio Balcao Dois').first()
        self.assertIsNotNone(it_ses1)
        self.assertIsNotNone(it_ses2)
        self.assertEqual(it_ses1.pagamento.forma, 'dinheiro')
        self.assertEqual(it_ses2.pagamento.forma, 'dinheiro')
        # O asterisco deve ter sido removido do nome
        self.assertNotIn('*', it_ses1.nome_intencao)
        self.assertEqual(it_ses1.nome_intencao, 'Jose Balcao Um')
