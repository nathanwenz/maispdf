import os
import subprocess
import streamlit as st
from docx import Document
from docxcompose.composer import Composer
from pypdf import PdfWriter

# Configuração da página
st.set_page_config(
    page_title="Mesclador de Documentos Word & PDF",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Mesclador de Documentos Word & PDF")
st.markdown("Junte múltiplos arquivos Word (.docx) ou PDF em um único arquivo de forma rápida e automática.")

tab1, tab2 = st.tabs(["📝 Mesclar Documentos Word (.docx)", "📕 Mesclar Arquivos PDF (.pdf)"])

# ---------------------------------------------------------
# TAB 1: MESCLAR WORD
# ---------------------------------------------------------
with tab1:
    st.header("Mesclar Arquivos Word")
    st.write("Selecione os arquivos .docx na ordem em que deseja juntá-los.")
    
    uploaded_word_files = st.file_uploader(
        "Escolha os arquivos Word (.docx)", 
        type=["docx"], 
        accept_multiple_files=True,
        key="word_uploader"
    )
    
    if uploaded_word_files:
        st.subheader(f"Total de arquivos selecionados: {len(uploaded_word_files)}")
        
        filenames = [f.name for f in uploaded_word_files]
        st.write("Ordem dos arquivos para fusão:")
        for idx, name in enumerate(filenames, start=1):
            st.text(f"{idx}. {name}")
            
        output_format = st.radio(
            "Formato de saída desejado:",
            options=["Apenas Word (.docx)", "Apenas PDF (.pdf)", "Ambos (Word e PDF)"],
            index=2
        )
        
        if st.button("🚀 Processar e Juntar Arquivos Word", type="primary"):
            with st.spinner("Processando e mesclando documentos..."):
                try:
                    # Salvar primeiro arquivo como base
                    first_file = uploaded_word_files[0]
                    base_doc = Document(first_file)
                    composer = Composer(base_doc)
                    
                    # Anexar os demais arquivos
                    for f in uploaded_word_files[1:]:
                        next_doc = Document(f)
                        composer.append(next_doc)
                        
                    merged_docx_path = "documento_final_mesclado.docx"
                    composer.save(merged_docx_path)
                    
                    st.success("✅ Arquivos Word mesclados com sucesso!")
                    
                    col1, col2 = st.columns(2)
                    
                    if output_format in ["Apenas Word (.docx)", "Ambos (Word e PDF)"]:
                        with open(merged_docx_path, "rb") as fp:
                            col1.download_button(
                                label="📥 Baixar Word Final (.docx)",
                                data=fp.read(),
                                file_name="documento_final_mesclado.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                            
                    if output_format in ["Apenas PDF (.pdf)", "Ambos (Word e PDF)"]:
                        # Converter usando LibreOffice
                        cmd = ["soffice", "--headless", "--convert-to", "pdf", merged_docx_path]
                        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                        
                        pdf_path = "documento_final_mesclado.pdf"
                        if os.path.exists(pdf_path):
                            with open(pdf_path, "rb") as fp:
                                col2.download_button(
                                    label="📥 Baixar PDF Final (.pdf)",
                                    data=fp.read(),
                                    file_name="documento_final_mesclado.pdf",
                                    mime="application/pdf"
                                )
                        else:
                            st.error("Erro na conversão para PDF pelo LibreOffice.")
                            
                except Exception as e:
                    st.error(f"Erro ao processar documentos Word: {str(e)}")

# ---------------------------------------------------------
# TAB 2: MESCLAR PDF
# ---------------------------------------------------------
with tab2:
    st.header("Mesclar Arquivos PDF")
    st.write("Selecione múltiplos PDFs soltos para juntar em um único PDF novo.")
    
    uploaded_pdf_files = st.file_uploader(
        "Escolha os arquivos PDF (.pdf)", 
        type=["pdf"], 
        accept_multiple_files=True,
        key="pdf_uploader"
    )
    
    if uploaded_pdf_files:
        st.subheader(f"Total de PDFs selecionados: {len(uploaded_pdf_files)}")
        
        pdf_names = [f.name for f in uploaded_pdf_files]
        st.write("Ordem de junção dos PDFs:")
        for idx, name in enumerate(pdf_names, start=1):
            st.text(f"{idx}. {name}")
            
        if st.button("🚀 Juntar Todos os PDFs", type="primary"):
            with st.spinner("Unificando PDFs..."):
                try:
                    writer = PdfWriter()
                    
                    for pdf_file in uploaded_pdf_files:
                        writer.append(pdf_file)
                        
                    output_pdf_path = "pdfs_unificados.pdf"
                    with open(output_pdf_path, "wb") as f_out:
                        writer.write(f_out)
                    writer.close()
                    
                    st.success("✅ PDFs unificados com sucesso!")
                    
                    with open(output_pdf_path, "rb") as fp:
                        st.download_button(
                            label="📥 Baixar PDF Unificado (.pdf)",
                            data=fp.read(),
                            file_name="pdfs_unificados.pdf",
                            mime="application/pdf"
                        )
                except Exception as e:
                    st.error(f"Erro ao juntar os arquivos PDF: {str(e)}")

# Rodapé com nota de privacidade
st.markdown("---")
st.caption("🔒 Seus arquivos são processados temporariamente na memória do servidor e descartados após o uso. Nenhuma cópia é armazenada.")
