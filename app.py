import streamlit as st
from docx import Document
from docx.shared import Pt
from io import BytesIO
from fpdf import FPDF
import os

# ==========================================================
# CONFIGURAÇÕES INICIAIS DA PÁGINA
# ==========================================================
st.set_page_config(
    page_title="Gerador de Propostas Técnicas",
    layout="wide"
)

st.title("📄 Gerador de Propostas Comerciais Técnicas")

# ==========================================================
# FUNÇÕES UTILITÁRIAS
# ==========================================================

def substituir_placeholders(doc, dados):
    """
    Substitui placeholders no formato {{PLACEHOLDER}}
    dentro do documento Word.
    """
    for p in doc.paragraphs:
        for chave, valor in dados.items():
            if f"{{{{{chave}}}}}" in p.text:
                p.text = p.text.replace(f"{{{{{chave}}}}}", valor)

    # Também substitui dentro de tabelas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for chave, valor in dados.items():
                    if f"{{{{{chave}}}}}" in cell.text:
                        cell.text = cell.text.replace(f"{{{{{chave}}}}}", valor)

    return doc


def gerar_docx(template_path, dados):
    """
    Carrega o template DOCX e aplica substituições.
    Retorna o arquivo em memória.
    """
    doc = Document(template_path)
    doc = substituir_placeholders(doc, dados)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


from fpdf import FPDF
from io import BytesIO

def gerar_pdf(dados):
    """
    Gera um PDF simples e retorna como BytesIO.
    """
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.cell(0, 10, "Proposta Comercial Técnica", ln=True, align="C")
    pdf.ln(5)

    for chave, valor in dados.items():
        pdf.multi_cell(0, 8, f"{chave.replace('_', ' ').title()}: {valor}")
        pdf.ln(2)

    # 🔥 AQUI ESTÁ A CORREÇÃO
    pdf_output = pdf.output(dest="S").encode("latin-1")

    buffer = BytesIO(pdf_output)
    buffer.seek(0)

    return buffer


# ==========================================================
# SIDEBAR - ENTRADAS DO USUÁRIO
# ==========================================================
st.sidebar.header("📝 Dados da Proposta")

nome_cliente = st.sidebar.text_input("Nome do Cliente")
titulo_projeto = st.sidebar.text_input("Título do Projeto")
valor_total = st.sidebar.text_input("Valor Total (R$)")
prazo_entrega = st.sidebar.text_input("Prazo de Entrega")
escopo_tecnico = st.sidebar.text_area("Escopo Técnico")

# Dicionário de dados para substituição
dados_proposta = {
    "NOME_CLIENTE": nome_cliente,
    "TITULO_PROJETO": titulo_projeto,
    "VALOR_TOTAL": valor_total,
    "PRAZO_ENTREGA": prazo_entrega,
    "ESCOPO_TECNICO": escopo_tecnico
}

# ==========================================================
# VISUALIZAÇÃO RESUMO
# ==========================================================
st.subheader("📋 Resumo da Proposta")

st.markdown(f"""
**Cliente:** {nome_cliente}  
**Projeto:** {titulo_projeto}  
**Valor Total:** R$ {valor_total}  
**Prazo:** {prazo_entrega}  
**Escopo Técnico:**  
{escopo_tecnico}
""")

st.divider()

# ==========================================================
# GERAÇÃO DO DOCUMENTO
# ==========================================================

TEMPLATE_PATH = "template.docx"

if os.path.exists(TEMPLATE_PATH):

    col1, col2 = st.columns(2)

    with col1:
        if st.button("📄 Gerar Proposta (DOCX)"):
            arquivo_docx = gerar_docx(TEMPLATE_PATH, dados_proposta)

            st.download_button(
                label="⬇️ Baixar Proposta DOCX",
                data=arquivo_docx,
                file_name=f"Proposta_{nome_cliente}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    with col2:
        if st.button("📄 Gerar Proposta (PDF)"):
            arquivo_pdf = gerar_pdf(dados_proposta)

            st.download_button(
                label="⬇️ Baixar Proposta PDF",
                data=arquivo_pdf,
                file_name=f"Proposta_{nome_cliente}.pdf",
                mime="application/pdf"
            )

else:
    st.error("Template 'template.docx' não encontrado na raiz do projeto.")
