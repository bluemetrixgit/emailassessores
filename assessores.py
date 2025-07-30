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
import os
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# Carregar .env do mesmo diretório
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
                    .str.strip()
                    .str.replace('.0', '', regex=False)
                    .str.extract(r'(\d+)')[0]  # pega apenas números
                    .fillna('')
                    .str.zfill(8)
                )
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

        st.write("Contas no controle:", controle['CONTA'].unique()[:10])
        st.write("Contas nas ordens:", ordens['CONTA'].unique()[:10])
        st.write("Contas nos acompanhamentos:", acompanhamento['CONTA'].unique()[:10])

       
    # Função que converte corretamente números no formato BR/US
    def to_float_safe(x):
        try:
            if pd.isnull(x): 
                return 0.0
            x = str(x).strip().replace(".", "").replace(",", ".")  # remove separador de milhar e ajusta decimal
            return float(x)
        except:
            return 0.0
        
        ordens['QT. EXECUTADA'] = ordens['QT. EXECUTADA'].apply(to_float_safe)
        ordens['PREÇO MÉDIO'] = ordens['PREÇO MÉDIO'].apply(to_float_safe)
        ordens['VALOR'] = (ordens['QT. EXECUTADA'] * ordens['PREÇO MÉDIO']).round(2)
        
        # Formatar como moeda
        ordens['VALOR'] = ordens['VALOR'].apply(
            lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            if pd.notnull(x) and x != '' else ''
        )


        colunas_operacionais = ['CONTA', 'DESCRIÇÃO', 'OPERAÇÃO', 'SITUAÇÃO', 'SOLICITADA', 'VALOR']

        # Se ordens vier vazio, cria um DataFrame vazio com as colunas corretas
        if ordens.empty:
            ordens = pd.DataFrame(columns=colunas_operacionais)
        else:
            for col in colunas_operacionais:
                if col not in ordens.columns:
                    ordens[col] = None
            ordens = ordens[colunas_operacionais]
        
        # Ajusta acompanhamento para as mesmas colunas
        for col in colunas_operacionais:
            if col not in acompanhamento.columns:
                acompanhamento[col] = None
        acompanhamento = acompanhamento[colunas_operacionais]
        
        # Junta ordens + acompanhamento
        movimentacoes = pd.concat([ordens, acompanhamento], ignore_index=True)
        
        # --- DEBUG VISUAL ---
        st.markdown("### DEBUG")
        st.write("Movimentações antes do merge (primeiras 20 linhas):")
        st.dataframe(movimentacoes.head(20))
        
        if 'CONTA' in movimentacoes.columns:
            st.write("Contas movimentações:", list(movimentacoes['CONTA'].unique()))
        else:
            st.write("Movimentações sem coluna CONTA!")
        
        if 'CONTA' in controle.columns:
            st.write("Contas controle:", list(controle['CONTA'].unique()))
        else:
            st.write("Controle sem coluna CONTA!")
        
        # Merge preservando apenas contas movimentadas
        base = pd.merge(movimentacoes, controle, on='CONTA', how='left', suffixes=('', '_DUP'))

        # Merge preservando todas as contas do controle
        base = pd.merge(movimentacoes, controle, on='CONTA', how='left', suffixes=('', '_DUP'))

        base = base.loc[:, ~base.columns.str.endswith('_DUP')]
        colunas_finais = ['CONTA', 'ASSESSOR', 'UF', 'OPERAÇÃO', 'DESCRIÇÃO', 'SITUAÇÃO', 'SOLICITADA', 'VALOR']
        return base[[col for col in colunas_finais if col in base.columns]]

    def gerar_pdf(self, assessor, data_dia, tabela):
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

        # --- Espaço para não colar no cabeçalho ---
        story.append(Spacer(1, 5*inch))

        # --- TABELA ---
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
                        val = float(str(val).replace("R$","").replace(".","").replace(",",".")) 
                        val = f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    except:
                        pass
                elif col == "SOLICITADA":
                    try:
                        val = pd.to_datetime(val).strftime("%d/%m/%Y")
                    except:
                        pass
                if col == "DESCRIÇÃO":
                    val = Paragraph(str(val), wrap_style)
                linha.append(val)
            dados.append(linha)

        larguras = [50, 80, 40, 80, 250, 80, 70, 80]  # Descrição menor, valor maior
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

                # --- Cabeçalho com logo + infos ---
        def cabecalho(canvas, doc):
            # Caminho absoluto da logo
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.jpg")
            page_width, page_height = landscape(letter)

            # Desenha a logo centralizada
            if os.path.exists(logo_path):
                logo_width = 700  # tamanho da logo
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

            # Título
            canvas.setFont("Helvetica-Bold", 14)
            canvas.drawCentredString(page_width / 2, y_position - 15, "Acompanhamento diário de operações")

            # Assessor e Data
            canvas.setFont("Helvetica", 10)
            canvas.drawCentredString(page_width / 2, y_position - 35, f"Assessor: {assessor}")
            canvas.drawCentredString(page_width / 2, y_position - 50, f"Data: {data_dia}")

        doc.build(story, onFirstPage=cabecalho)
        return nome_pdf


    def enviar_email(self, assessor, destinatario, nome_pdf, data_dia):
        remetente = "david.alves@bluemetrix.com.br"
        assunto = f"Acompanhamento diário de operações - {assessor} - {data_dia}"
    
        # Cria a mensagem
        msg = MIMEMultipart()
        msg['From'] = remetente
        msg['To'] = destinatario
        msg['Subject'] = assunto
    
        # Corpo HTML com assinatura
        mensagem_html = f"""
        <html>
          <body style="font-family: Arial, sans-serif; color: #333;">
            <p>Olá {assessor},</p>
            <p>Segue em anexo o relatório de operações do dia {data_dia}.</p>
            <p>Qualquer dúvida, estamos à disposição.</p>
            <br>
            <img src="cid:assinatura" alt="Assinatura Bluemetrix" style="width:500px; height:auto;">
          </body>
        </html>
        """
        msg.attach(MIMEText(mensagem_html, 'html'))
    
        # Adiciona a imagem da assinatura
        assinatura_path = os.path.join(os.path.dirname(__file__), "Assinatura David.jpg")
        with open(assinatura_path, 'rb') as f:
            img = MIMEImage(f.read(), _subtype='jpeg')
            img.add_header('Content-ID', '<assinatura>')
            img.add_header('Content-Disposition', 'inline', filename="Assinatura.jpg")
            msg.attach(img)
            
        # Anexa o PDF
        with open(nome_pdf, 'rb') as f:
            attach = MIMEApplication(f.read(), _subtype='pdf')
            attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(nome_pdf))
            msg.attach(attach)
    
        # Envia via SMTP
        with smtplib.SMTP('smtp.kinghost.net', 587) as server:
            server.starttls()
            server.login(remetente, os.getenv("EMAIL_PASSWORD"))
            server.sendmail(remetente, destinatario, msg.as_string())
        
            return True
