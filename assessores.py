
import pandas as pd
import streamlit as st
import datetime
from reportlab.platypus import SimpleDocTemplate, Table, Paragraph, KeepTogether, Spacer
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.header import Header
from email.utils import formataddr
import os
from dotenv import load_dotenv

# Carregar .env do mesmo diretório
dotenv_path = os.path.join(os.path.dirname(__file__), ".env")
load_dotenv(dotenv_path=dotenv_path)

# Mantido para compatibilidade, mas não é mais utilizado diretamente
dia_e_hora = datetime.datetime.now() - datetime.timedelta(days=1)

class Comercial:
    def __init__(self):
        pass

    def tratando_dados(self, ordens, acompanhamento, controle):
        for df in [ordens, acompanhamento, controle]:
            df.columns = df.columns.str.upper()
            if 'CONTA' in df.columns:
                df['CONTA'] = (
                    df['CONTA']
                    .astype(str)
                    .str.extract(r'(\d+)')[0]   # pega só os números
                    .fillna('')
                    .str.strip()
                    .str.zfill(9)               # completa sempre com zeros à esquerda
                )
            
            df['CONTA'] = df['CONTA'].apply(lambda x: x.zfill(9)[-9:])

        if 'OPERACAO' in acompanhamento.columns:
            acompanhamento = acompanhamento.rename(columns={'OPERACAO': 'OPERAÇÃO'})
        if 'DESCRICAO' in acompanhamento.columns:
            acompanhamento = acompanhamento.rename(columns={'DESCRICAO': 'DESCRIÇÃO'})
        if 'SITUACAO' in acompanhamento.columns:
            acompanhamento = acompanhamento.rename(columns={'SITUACAO': 'SITUAÇÃO'})
        ordens = ordens.rename(columns={
            'DIREÇÃO': 'OPERAÇÃO',
            'ATIVO': 'DESCRIÇÃO',
            'STATUS': 'SITUAÇÃO',
            'DATA/HORA': 'SOLICITADA'})
        if 'SITUAÇÃO' in controle.columns:
            controle = controle.drop(columns=['SITUAÇÃO'])
        if 'VALOR FINANCEIRO' in ordens.columns:
            ordens['VALOR'] = ordens['VALOR FINANCEIRO']

        def to_float_safe(x):
            x = str(x).strip()
            if x == '' or x.lower() == 'nan':
                return 0.0
            if ',' in x and '.' in x and x.find(',') > x.find('.'):
                x = x.replace('.', '').replace(',', '.')
            elif ',' in x:
                x = x.replace(',', '.')
            return float(x)

        if not ordens.empty:
            if 'QT. EXECUTADA' in ordens.columns and 'PREÇO MÉDIO' in ordens.columns:
                ordens['QT. EXECUTADA'] = ordens['QT. EXECUTADA'].apply(to_float_safe)
                ordens['PREÇO MÉDIO'] = ordens['PREÇO MÉDIO'].apply(to_float_safe)
                ordens['VALOR'] = (ordens['QT. EXECUTADA'] * ordens['PREÇO MÉDIO']).round(2)
                ordens['VALOR'] = ordens['VALOR'].apply(
                    lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    if pd.notnull(x) and x != '' else ''
                )

        colunas_operacionais = ['CONTA', 'DESCRIÇÃO', 'OPERAÇÃO', 'SITUAÇÃO', 'SOLICITADA', 'VALOR']
        ordens = ordens[[col for col in colunas_operacionais if col in ordens.columns]]
        acompanhamento = acompanhamento[[col for col in colunas_operacionais if col in acompanhamento.columns]]

        for df in [ordens, acompanhamento]:
            if 'SOLICITADA' in df.columns:
                # Converte para datetime garantindo dayfirst (Brasil)
                df['SOLICITADA'] = pd.to_datetime(df['SOLICITADA'], errors='coerce', dayfirst=True)
                # Converte de volta para string no formato brasileiro
                df['SOLICITADA'] = df['SOLICITADA'].dt.strftime("%d/%m/%Y %H:%M:%S")


        movimentacoes = pd.concat([ordens, acompanhamento], ignore_index=True)
        base = pd.merge(controle, movimentacoes, on='CONTA', how='inner', suffixes=('', '_DUP'))
        base = base.loc[:, ~base.columns.str.endswith('_DUP')]
        base = pd.merge(controle, movimentacoes, on='CONTA', how='inner', suffixes=('', '_DUP'))
        colunas_finais = ['CONTA', 'ASSESSOR', 'UF', 'OPERAÇÃO', 'DESCRIÇÃO', 'SITUAÇÃO', 'SOLICITADA', 'VALOR']
        return base[[col for col in colunas_finais if col in base.columns]]

    
    def gerar_pdf(self, assessor, data_ini, data_fim, tabela):
        pasta_pdfs = os.path.join(os.path.dirname(__file__), "pdfs")
        os.makedirs(pasta_pdfs, exist_ok=True)
        nome_pdf = os.path.join(pasta_pdfs, f"Relatorio_{assessor.replace(' ', '_')}.pdf")

        doc = SimpleDocTemplate(
            nome_pdf,
            pagesize=landscape(letter),
            rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25
        )
        story = []
        styles = getSampleStyleSheet()
        wrap_style = ParagraphStyle('wrap', fontSize=5.5, leading=6)

        if tabela.empty:
            story.append(Paragraph("Nenhuma operação encontrada.", styles['Normal']))
            doc.build(story)
            return None

        story.append(Spacer(1, 5*inch))

        colunas = [col for col in tabela.columns if col in ["CONTA", "ASSESSOR", "UF", "OPERAÇÃO", "DESCRIÇÃO", "SITUAÇÃO", "SOLICITADA", "VALOR"]]
        dados = [colunas]
        for _, row in tabela.iterrows():
            linha = []
            for col in colunas:
                val = row.get(col, "")
                if pd.isna(val):
                    val = ""
                elif col == "VALOR":
                    try:
                        val = float(str(val).replace("R$", "").replace(".", "").replace(",", "."))
                        val = f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    except:
                        pass
                elif col == "SOLICITADA":
                    try:
                        val = pd.to_datetime(val, dayfirst=True).strftime("%d/%m/%Y")
                    except:
                        pass
                if col == "DESCRIÇÃO":
                    val = Paragraph(str(val), wrap_style)
                linha.append(val)
            dados.append(linha)

        larguras = [50, 80, 40, 80, 250, 80, 70, 80]
        tabela_pdf = Table(dados, repeatRows=1, colWidths=larguras, hAlign='LEFT')
        tabela_pdf.setStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
        ])
        story.append(KeepTogether(tabela_pdf))

        def cabecalho(canvas, doc, _data_ini=data_ini, _data_fim=data_fim):
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.jpg")
            page_width, page_height = landscape(letter)
            if os.path.exists(logo_path):
                logo_width = 700
                logo_height = 300
                x_position = (page_width - logo_width) / 2
                y_position = page_height - 300
                canvas.drawImage(
                    logo_path,
                    x_position,
                    y_position,
                    width=logo_width,
                    height=logo_height,
                    preserveAspectRatio=True,
                    mask='auto'
                )
            else:
                y_position = page_height - 50

            canvas.setFont("Helvetica-Bold", 14)
            canvas.drawCentredString(page_width / 2, y_position - 15, "Acompanhamento de operações")
            canvas.setFont("Helvetica", 10)
            canvas.drawCentredString(page_width / 2, y_position - 35, f"Assessor: {assessor}")

            try:
                di = pd.to_datetime(_data_ini).strftime("%d/%m/%Y")
                df = pd.to_datetime(_data_fim).strftime("%d/%m/%Y")
            except Exception:
                di = str(_data_ini)
                df = str(_data_fim)

            linha_data = f"Data: {di}" if di == df else f"Período: {di} a {df}"
            canvas.drawCentredString(page_width / 2, y_position - 50, linha_data)

        doc.build(story, onFirstPage=cabecalho)
        return nome_pdf

    def enviar_email(self, assessor, destinatario, nome_pdf, data_ini, data_fim):
        remetente = st.secrets["EMAIL_USER"]

        try:
            di = pd.to_datetime(data_ini).strftime("%d/%m/%Y")
            df = pd.to_datetime(data_fim).strftime("%d/%m/%Y")
        except Exception:
            di = str(data_ini)
            df = str(data_fim)

        if di == df:
            assunto = f"Acompanhamento diário de operações - {assessor} - {di}"
            corpo_data = f"do dia {di}"
        else:
            assunto = f"Acompanhamento de operações - {assessor} - {di} a {df}"
            corpo_data = f"do período {di} a {df}"

        msg = MIMEMultipart()
        msg['From'] = formataddr((str(Header("Middle Office Bluemetrix", "utf-8")), remetente))
        msg['To'] = formataddr((str(Header(assessor, 'utf-8')), destinatario))
        msg['Subject'] = Header(assunto, "utf-8")

        mensagem_html = f"""
        <html>
          <body style="font-family: Arial, sans-serif; color: #333;">
            <p>Olá {assessor},</p>
            <p>Segue em anexo o relatório de operações {corpo_data}.</p>
            <p>Qualquer dúvida, estamos à disposição.</p>
            <br>
            <img src="cid:assinatura" alt="Assinatura Bluemetrix" style="width:500px; height:auto;">
          </body>
        </html>
        """
        msg.attach(MIMEText(mensagem_html, 'html', 'utf-8'))

        assinatura_path = os.path.join(os.path.dirname(__file__), "Assinatura David.jpg")
        if os.path.exists(assinatura_path):
            with open(assinatura_path, 'rb') as f:
                img = MIMEImage(f.read(), _subtype='jpeg')
                img.add_header('Content-ID', '<assinatura>')
                img.add_header('Content-Disposition', 'inline', filename="Assinatura.jpg")
                msg.attach(img)

        if nome_pdf and os.path.exists(nome_pdf):
            with open(nome_pdf, 'rb') as f:
                attach = MIMEApplication(f.read(), _subtype='pdf')
                attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(nome_pdf))
                msg.attach(attach)

        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as server:
            server.starttls()
            server.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASSWORD"])
            server.sendmail(remetente, destinatario, msg.as_string())
        return True
