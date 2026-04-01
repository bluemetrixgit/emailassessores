
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

dotenv_path = os.path.join(os.path.dirname(__file__), ".env")
load_dotenv(dotenv_path=dotenv_path)

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
                    .str.extract(r'(\d+)')[0]
                    .fillna('')
                    .str.strip()
                    .str.zfill(9)
                )

                df['CONTA'] = df['CONTA'].apply(lambda x: str(x).zfill(9)[-9:])

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
            'DATA/HORA': 'SOLICITADA'
        })

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

        # --- CORREÇÃO PRINCIPAL (dt seguro) ---
        if 'SOLICITADA' in ordens.columns:
            s = ordens['SOLICITADA']

            if not isinstance(s, pd.Series):
                s = pd.Series(s)

            s = s.astype(str).str.strip()

            dt = pd.to_datetime(s, errors='coerce', dayfirst=True)

            mask = dt.isna()
            if mask.any():
                dt.loc[mask] = pd.to_datetime(s[mask], errors='coerce')

            ordens['SOLICITADA'] = dt.dt.strftime("%d/%m/%Y")

        if 'SOLICITADA' in acompanhamento.columns:
            acompanhamento['SOLICITADA'] = pd.to_datetime(
                acompanhamento['SOLICITADA'], errors='coerce', dayfirst=True
            ).dt.strftime("%d/%m/%Y")

        movimentacoes = pd.concat([ordens, acompanhamento], ignore_index=True)

        base = pd.merge(controle, movimentacoes, on='CONTA', how='inner')

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

        story.append(Spacer(1, 5 * inch))

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

        tabela_pdf = Table(dados, repeatRows=1, hAlign='LEFT')
        tabela_pdf.setStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
        ])

        story.append(KeepTogether(tabela_pdf))
        doc.build(story)

        return nome_pdf

    def enviar_email(self, assessor, destinatario, nome_pdf, data_ini, data_fim):
        remetente = st.secrets["EMAIL_USER"]

        msg = MIMEMultipart()
        msg['From'] = remetente
        msg['To'] = destinatario
        msg['Subject'] = f"Relatório - {assessor}"

        msg.attach(MIMEText("Segue relatório em anexo.", 'plain'))

        if nome_pdf and os.path.exists(nome_pdf):
            with open(nome_pdf, 'rb') as f:
                attach = MIMEApplication(f.read(), _subtype='pdf')
                attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(nome_pdf))
                msg.attach(attach)

        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASSWORD"])
            server.sendmail(remetente, destinatario, msg.as_string())

        return True

        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as server:
            server.starttls()
            server.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASSWORD"])
            server.sendmail(remetente, destinatario, msg.as_string())
        return True
