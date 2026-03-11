import streamlit as st
from docx import Document
from docx.shared import Inches
from io import BytesIO
import tempfile
import os

# ==========================================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================================
st.set_page_config(page_title="Gerador de Propostas", layout="wide")

# Exibir logo da empresa contratada no topo da interface
st.image("LOGO DGCE.png", width=200)  # logo fixa no projeto

st.title("📄 Gerador de Propostas Comerciais Técnicas")

# ==========================================================
# FUNÇÕES
# ==========================================================

def substituir_placeholders(doc, dados):
    """
    Substitui placeholders no formato {{PLACEHOLDER}}
    """
    for p in doc.paragraphs:
        if "{{LOGO}}" in p.text:
            p.clear()  # remove texto do placeholder
            run = p.add_run()
            run.add_picture("LOGO DGCE.png", width=Inches(2))
        else:
            for chave, valor in dados.items():
                if f"{{{{{chave}}}}}" in p.text:
                    p.text = p.text.replace(f"{{{{{chave}}}}}", valor)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for chave, valor in dados.items():
                    if f"{{{{{chave}}}}}" in cell.text:
                        cell.text = cell.text.replace(f"{{{{{chave}}}}}", valor)

    return doc


def gerar_docx(template_file, dados):
    doc = Document(template_file)
    doc = substituir_placeholders(doc, dados)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def converter_docx_para_pdf(docx_bytes):
    """
    Conversão real mantendo layout.
    Funciona no Streamlit Cloud usando docx2pdf local fallback.
    """
    try:
        from docx2pdf import convert

        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = os.path.join(tmpdir, "arquivo.docx")
            pdf_path = os.path.join(tmpdir, "arquivo.pdf")

            with open(docx_path, "wb") as f:
                f.write(docx_bytes.getbuffer())

            convert(docx_path, pdf_path)

            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()

        return BytesIO(pdf_bytes)

    except Exception as e:
        st.error("Conversão para PDF não disponível no ambiente atual.")
        return None


# ==========================================================
# SIDEBAR - ENTRADAS
# ==========================================================
st.sidebar.header("📝 Dados da Proposta")

nome_cliente = st.sidebar.text_input("Nome do Cliente")
titulo_projeto = st.sidebar.text_input("Título do Projeto")
valor_total = st.sidebar.text_input("Valor Total (R$)")
prazo_entrega = st.sidebar.text_input("Prazo de Entrega")
escopo_tecnico = st.sidebar.text_area("Escopo Técnico")

template_file = st.sidebar.file_uploader(
    "Upload do Template (.docx)",
    type=["docx"]
)

# ==========================================================
# VALIDAÇÃO
# ==========================================================
campos_preenchidos = all([
    nome_cliente,
    titulo_projeto,
    valor_total,
    prazo_entrega,
    escopo_tecnico,
    template_file
])

dados_proposta = {
    "NOME_CLIENTE": nome_cliente,
    "TITULO_PROJETO": titulo_projeto,
    "VALOR_TOTAL": valor_total,
    "PRAZO_ENTREGA": prazo_entrega,
    "ESCOPO_TECNICO": escopo_tecnico
}

# ==========================================================
# RESUMO
# ==========================================================
st.subheader("📋 Resumo da Proposta")

st.markdown(f"""
**Cliente:** {nome_cliente}  
**Projeto:** {titulo_projeto}  
**Valor:** R$ {valor_total}  
**Prazo:** {prazo_entrega}  

**Escopo Técnico:**  
{escopo_tecnico}
""")

st.divider()

# ==========================================================
# GERAÇÃO
# ==========================================================
if campos_preenchidos:
    if st.button("🚀 Gerar Proposta"):
        arquivo_docx = gerar_docx(template_file, dados_proposta)

        col1, col2 = st.columns(2)

        with col1:
            st.download_button(
                label="⬇️ Baixar DOCX",
                data=arquivo_docx,
                file_name=f"Proposta_{nome_cliente}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        with col2:
            pdf_file = converter_docx_para_pdf(arquivo_docx)

            if pdf_file:
                st.download_button(
                    label="⬇️ Baixar PDF",
                    data=pdf_file,
                    file_name=f"Proposta_{nome_cliente}.pdf",
                    mime="application/pdf"
                )
else:
    st.warning("Preencha todos os campos e envie o template para gerar a proposta.")

