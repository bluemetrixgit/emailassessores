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

# Garante que a pasta de PDFs existe
os.makedirs("pdfs", exist_ok=True)

BASE_DIR = os.getcwd()

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

ordens = pd.read_csv("ordens.csv", encoding="utf-8", sep=",")
acompanhamentos = pd.read_csv("acompanhamento_de_operacoes.csv", encoding="utf-8", sep=",")
emails = pd.read_csv("emails.csv", encoding="utf-8", sep=",")

# Exibir prévia
st.subheader("Prévia das Operações")
st.dataframe(arquivo_final)

# Criar pasta de PDFs se não existir
os.makedirs(os.path.join(BASE_DIR, 'pdfs'), exist_ok=True)

# --- Botão único: Gera e envia PDFs ---
if st.button("📄 Gerar e Enviar Relatórios"):
    df_emails = ler_csv_seguro(os.path.join(BASE_DIR, "emails.csv"))
    consolidados = df_emails[df_emails["CONSOLIDADO"].str.upper() == "SIM"]["NOME"].str.strip().tolist()
    assessores_unicos = arquivo_final['ASSESSOR'].dropna().unique()
    destinatarios = list(assessores_unicos) + consolidados

    for destinatario in destinatarios:
        try:
            tabela = arquivo_final if destinatario in consolidados else arquivo_final[arquivo_final['ASSESSOR'] == destinatario]
            nome_pdf = comercial.gerar_pdf(destinatario, dia_e_hora.strftime('%d/%m/%Y'), tabela)

            if nome_pdf and os.path.exists(nome_pdf):
                email_registro = df_emails[df_emails['NOME'].str.strip().str.upper() == destinatario.strip().upper()]['EMAIL']
                if email_registro.empty:
                    st.warning(f"Destinatário {destinatario} não possui e-mail na planilha. Pulando...")
                    continue
                email = email_registro.values[0]
                comercial.enviar_email(destinatario, email, nome_pdf, dia_e_hora.strftime('%d/%m/%Y'))
                st.success(f"Relatório de {destinatario} gerado e enviado!")
            else:
                st.warning(f"Não foi possível gerar o PDF para {destinatario}.")
        except Exception as e:
            st.error(f"Erro ao processar {destinatario}: {e}")
