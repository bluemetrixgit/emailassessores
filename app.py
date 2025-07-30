import streamlit as st
import pandas as pd
import os
from assessores import Comercial, dia_e_hora

# Configuração da página
st.set_page_config(page_title="Envio de Operações", layout="wide")
st.title("Envio de Operações - Bluemetrix")

comercial = Comercial()

# Função auxiliar para leitura segura de CSV
def ler_csv_seguro(caminho):
    try:
        return pd.read_csv(caminho, encoding='utf-8')
    except UnicodeDecodeError:
        return pd.read_csv(caminho, encoding='latin1')

# Carregar arquivos de entrada
ordens = ler_csv_seguro('ordens.csv')
acompanhamentos = ler_csv_seguro('acompanhamento_de_operacoes.csv')
controle_excel = pd.ExcelFile('Controle de Contratos - Atualizado 2025.xlsx')
controle = controle_excel.parse('BTG', header=1)
arquivo_final = comercial.tratando_dados(ordens, acompanhamentos, controle)

# Exibir prévia
st.subheader("Prévia das Operações")
st.dataframe(arquivo_final)

# Criar pasta de PDFs se não existir
os.makedirs('pdfs', exist_ok=True)

# --- Botão único: Gera e envia PDFs ---
if st.button("📄 Gerar e Enviar Relatórios"):
    # Carrega lista de e-mails e identifica consolidados
    df_emails = pd.read_csv("emails.csv", encoding="utf-8")
    consolidados = df_emails[df_emails["CONSOLIDADO"].str.upper() == "SIM"]["NOME"].str.strip().tolist()
    assessores_unicos = arquivo_final['ASSESSOR'].dropna().unique()
    destinatarios = list(assessores_unicos) + consolidados

    for destinatario in destinatarios:
        try:
            # Se for consolidado, usa o arquivo completo. Caso contrário, filtra pelo assessor
            tabela = arquivo_final if destinatario in consolidados else arquivo_final[arquivo_final['ASSESSOR'] == destinatario]

            # Gera o PDF
            nome_pdf = comercial.gerar_pdf(destinatario, dia_e_hora.strftime('%d/%m/%Y'), tabela)

            if nome_pdf and os.path.exists(nome_pdf):
                # Verifica se o destinatário existe no CSV (ignora se não existir)
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
