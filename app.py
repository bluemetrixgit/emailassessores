import streamlit as st
import pandas as pd
import os
from assessores import Comercial, dia_e_hora
import chardet

# Diretório base para todos os arquivos
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Configuração da página
st.set_page_config(page_title="Envio de Operações", layout="wide")
st.title("📈 Envio de Operações - Bluemetrix")

comercial = Comercial()

# Função para detectar encoding antes de ler
def detectar_encoding(caminho):
    with open(caminho, 'rb') as f:
        resultado = chardet.detect(f.read())
    return resultado['encoding']

# Função segura para carregar CSV
def ler_csv_seguro(caminho, sep=","):
    try:
        encoding_detectado = detectar_encoding(caminho)
        return pd.read_csv(caminho, encoding=encoding_detectado, sep=sep)
    except Exception:
        return pd.read_csv(caminho, encoding='latin1', sep=sep)

# Lendo arquivos
ordens = ler_csv_seguro(os.path.join(BASE_DIR, 'ordens.csv'))
acompanhamentos = ler_csv_seguro(os.path.join(BASE_DIR, 'acompanhamento_de_operacoes.csv'))
emails = ler_csv_seguro(os.path.join(BASE_DIR, 'emails.csv'))

# Excel
controle_excel = pd.ExcelFile(os.path.join(BASE_DIR, 'Controle de Contratos - Atualizado 2025.xlsx'))
try:
    controle = controle_excel.parse('BTG', header=1)
except:
    controle = controle_excel.parse(controle_excel.sheet_names[0], header=1)

arquivo_final = comercial.tratando_dados(ordens, acompanhamentos, controle)

# Exibir prévia
st.subheader("Prévia das Operações")
st.dataframe(arquivo_final)

# Criar pasta de PDFs se não existir
os.makedirs(os.path.join(BASE_DIR, 'pdfs'), exist_ok=True)

# --- Botão único: Gera e envia PDFs ---
if st.button("Gerar e Enviar Relatórios"):
    assessores_unicos = arquivo_final['ASSESSOR'].dropna().unique()
    for destinatario in assessores_unicos:
        tabela = arquivo_final if destinatario in consolidados else arquivo_final[arquivo_final['ASSESSOR'] == destinatario]
        if tabela.empty:
            st.warning(f"Destinatário {destinatario} não possui dados. Pulando...")
            continue

        try:
            st.write(f"➡️ Gerando PDF para {destinatario}...")
            comercial.gerar_pdf(destinatario, data_hoje, tabela)
            st.write("✅ PDF gerado com sucesso.")

            st.write(f"➡️ Conectando ao servidor SMTP para {destinatario}...")
            with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as server:
                st.write("➡️ Iniciando conexão segura...")
                server.starttls()
                st.write("➡️ Fazendo login...")
                server.login(remetente, os.getenv("EMAIL_PASSWORD"))
                st.write("➡️ Enviando e-mail...")
                server.sendmail(remetente, destinatario, msg.as_string())
                st.success(f"✅ E-mail enviado para {destinatario}.")
        except Exception as e:
            st.error(f"❌ Erro ao processar {destinatario}: {e}")
