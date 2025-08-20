
import streamlit as st
import pandas as pd
import os
import smtplib
from assessores import Comercial
import chardet
from datetime import date, timedelta

# Diretório base
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Configuração da página
st.set_page_config(page_title="Envio de Operações", layout="wide")
st.title("📈 Envio de Operações - Bluemetrix")

comercial = Comercial()

# Função para detectar encoding
def detectar_encoding(caminho):
    with open(caminho, 'rb') as f:
        resultado = chardet.detect(f.read())
    return resultado['encoding']

# Função segura para ler CSV
def ler_csv_seguro(caminho, sep=','):
    try:
        encoding_detectado = detectar_encoding(caminho)
        return pd.read_csv(caminho, encoding=encoding_detectado, sep=sep)
    except Exception:
        return pd.read_csv(caminho, encoding='latin1', sep=sep)

# Ler arquivos
ordens = ler_csv_seguro(os.path.join(BASE_DIR, 'ordens.csv'))
acompanhamentos = ler_csv_seguro(os.path.join(BASE_DIR, 'acompanhamento_de_operacoes.csv'))
emails = ler_csv_seguro(os.path.join(BASE_DIR, 'emails.csv'))

# Ler Excel de controle
controle_excel = pd.ExcelFile(os.path.join(BASE_DIR, 'Controle de Contratos - Atualizado 2025.xlsx'))
try:
    controle = controle_excel.parse('BTG', header=1)
except:
    controle = controle_excel.parse(controle_excel.sheet_names[0], header=1)

# Tratar dados
arquivo_final = comercial.tratando_dados(ordens, acompanhamentos, controle)

# ---------- Seletor de período ----------
def ultimo_dia_util(d: date) -> date:
    while d.weekday() >= 5:  # 5=sábado, 6=domingo
        d -= timedelta(days=1)
    return d

hoje = date.today()
default_fim = ultimo_dia_util(hoje)

data_ini, data_fim = st.date_input(
    "Período do relatório (início e fim)",
    value=(default_fim, default_fim),
    format="DD/MM/YYYY"
)

if data_ini > data_fim:
    data_ini, data_fim = data_fim, data_ini

st.write("Ordens contas únicas:", ordens['CONTA'].unique()[:10])
st.write("Acompanhamento contas únicas:", acompanhamentos['CONTA'].unique()[:10])
st.write("Controle contas únicas:", controle['CONTA'].unique()[:10])

# Cria coluna auxiliar de data (apenas a data) a partir de SOLICITADA
arquivo_final['_DATA'] = pd.to_datetime(
    arquivo_final['SOLICITADA'], dayfirst=True, errors='coerce'
).dt.date

# Filtra pelo período selecionado
arquivo_filtrado = arquivo_final.loc[
    (arquivo_final['_DATA'] >= data_ini) & (arquivo_final['_DATA'] <= data_fim)
].copy()

# Prévia
st.subheader("Prévia das Operações")
st.dataframe(arquivo_filtrado.drop(columns=['_DATA'], errors='ignore'))

# Criar pasta PDFs
os.makedirs(os.path.join(BASE_DIR, 'pdfs'), exist_ok=True)

# Botão para gerar e enviar relatórios
if st.button("Gerar e Enviar Relatórios"):
    # Mantém fluxo original: consolidados por NOME
    consolidados = emails.loc[emails['CONSOLIDADO'].str.upper() == 'SIM', 'NOME'].str.strip().tolist()
    assessores_unicos = arquivo_filtrado['ASSESSOR'].dropna().unique()
    destinatarios = list(assessores_unicos) + consolidados

    for destinatario in destinatarios:
        # Busca o e-mail correspondente ao nome
        email_destinatario = emails.loc[
            emails['NOME'].str.strip().str.upper() == str(destinatario).strip().upper(),
            'EMAIL'
        ].values

        if len(email_destinatario) == 0:
            st.warning(f"Não foi encontrado e-mail para {destinatario}. Pulando...")
            continue
        email_destinatario = email_destinatario[0]

        tabela = arquivo_filtrado if destinatario in consolidados else arquivo_filtrado[arquivo_filtrado['ASSESSOR'] == destinatario]
        if tabela.empty:
            st.warning(f"Destinatário {destinatario} não possui dados no período. Pulando...")
            continue

        try:
            st.write(f"➡️ Gerando PDF para {destinatario}...")
            nome_pdf = comercial.gerar_pdf(destinatario, data_ini, data_fim, tabela)  # agora com período
            st.write("✅ PDF gerado com sucesso.")

            st.write(f"➡️ Enviando e-mail para {destinatario}...")
            comercial.enviar_email(destinatario, email_destinatario, nome_pdf, data_ini, data_fim)
            st.success(f"✅ E-mail enviado para {destinatario}.")
        except Exception as e:
            st.error(f"❌ Erro ao processar {destinatario}: {e}")
