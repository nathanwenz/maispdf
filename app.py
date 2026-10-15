import os
import subprocess
import tempfile
import streamlit as st
from docx import Document
from docxcompose.composer import Composer
from pypdf import PdfWriter
from pdf2docx import Converter

# Configuração da página
st.set_page_config(
    page_title="+PDF - Converta e Mescle Documentos",
    page_icon="📄",
    layout="wide"
)

# ---------------------------------------------------------
# MARCA D'ÁGUA FIXA NO CANTO INFERIOR DIREITO
# ---------------------------------------------------------
st.markdown(
    """
    <style>
    .watermark {
        position: fixed;
        bottom: 12px;
        right: 20px;
        font-size: 13px;
        font-weight: bold;
        color: #666666;
        background-color: rgba(240, 242, 246, 0.85);
        padding: 6px 14px;
        border-radius: 6px;
        border: 1px solid rgba(200, 200, 200, 0.5);
        z-index: 999999;
        pointer-events: none;
        font-family: 'Source Sans Pro', sans-serif;
        text-align: right;
        line-height: 1.25;
    }
    </style>
    <div class="watermark">
        developer;<br>
        Nathan Wenzel
    </div>
    """,
    unsafe_allow_html=True
)

# ---------------------------------------------------------
# FOTO DE CAPA / BANNER DO SITE (90% DE LARGURA)
# ---------------------------------------------------------
banner_path = "banner.png" if os.path.exists("banner.png") else ("banner.jpg" if os.path.exists("banner.jpg") else None)
if banner_path:
    col_b1, col_b2, col_b3 = st.columns([0.05, 0.90, 0.05])
    with col_b2:
        st.image(banner_path, use_container_width=True)

# Título Principal do Aplicativo
st.title("📄 +PDF - Converta ou faça a junção de seus documentos (Word & PDF)")
st.markdown("Junte múltiplos documentos ou converta formatos entre Word e PDF de forma rápida e gratuita.")

# Criar as abas do aplicativo
tab1, tab2, tab3 = st.tabs([
    "📝 Mesclar Word (.docx)", 
    "📕 Mesclar PDF (.pdf)", 
    "🔄 Converter (Word ↔ PDF)"
])

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
            index=2,
            key="word_out_format"
        )
        
        if st.button("🚀 Processar e Juntar Arquivos Word", type="primary", key="btn_merge_word"):
            with st.spinner("Processando e mesclando documentos..."):
                try:
                    first_file = uploaded_word_files[0]
                    base_doc = Document(first_file)
                    composer = Composer(base_doc)
                    
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
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key="btn_dl_docx"
                            )
                            
                    if output_format in ["Apenas PDF (.pdf)", "Ambos (Word e PDF)"]:
                        cmd = ["soffice", "--headless", "--convert-to", "pdf", merged_docx_path]
                        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                        
                        pdf_path = "documento_final_mesclado.pdf"
                        if os.path.exists(pdf_path):
                            with open(pdf_path, "rb") as fp:
                                col2.download_button(
                                    label="📥 Baixar PDF Final (.pdf)",
                                    data=fp.read(),
                                    file_name="documento_final_mesclado.pdf",
                                    mime="application/pdf",
                                    key="btn_dl_pdf_from_word"
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
            
        if st.button("🚀 Juntar Todos os PDFs", type="primary", key="btn_merge_pdf"):
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
                            mime="application/pdf",
                            key="btn_dl_merged_pdf"
                        )
                except Exception as e:
                    st.error(f"Erro ao juntar os arquivos PDF: {str(e)}")

# ---------------------------------------------------------
# TAB 3: CONVERTER (WORD <-> PDF)
# ---------------------------------------------------------
with tab3:
    st.header("Converter Formatos (Word ↔ PDF)")
    st.write("Transforme arquivos Word em PDF ou converta PDFs para arquivos Word editáveis.")
    
    conversion_type = st.radio(
        "Escolha o tipo de conversão:",
        options=["Word (.docx) ➡️ PDF (.pdf)", "PDF (.pdf) ➡️ Word (.docx)"],
        horizontal=True,
        key="conversion_type"
    )
    
    st.divider()
    
    # OPÇÃO 1: WORD PARA PDF
    if conversion_type == "Word (.docx) ➡️ PDF (.pdf)":
        st.subheader("Converter Word para PDF")
        word_to_convert = st.file_uploader(
            "Selecione um arquivo Word (.docx)", 
            type=["docx"], 
            accept_multiple_files=False,
            key="w2p_uploader"
        )
        
        if word_to_convert:
            if st.button("🚀 Converter para PDF", type="primary", key="btn_w2p"):
                with st.spinner("Convertendo Word para PDF..."):
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                            tmp_docx.write(word_to_convert.getvalue())
                            tmp_docx_path = tmp_docx.name
                            
                        cmd = ["soffice", "--headless", "--convert-to", "pdf", tmp_docx_path, "--outdir", tempfile.gettempdir()]
                        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                        
                        base_name, _ = os.path.splitext(tmp_docx_path)
                        expected_pdf_path = base_name + ".pdf"
                        
                        if os.path.exists(expected_pdf_path):
                            st.success("✅ Conversão para PDF concluída!")
                            with open(expected_pdf_path, "rb") as fp:
                                original_base_name, _ = os.path.splitext(word_to_convert.name)
                                st.download_button(
                                    label="📥 Baixar PDF Convertido",
                                    data=fp.read(),
                                    file_name=f"{original_base_name}.pdf",
                                    mime="application/pdf",
                                    key="btn_dl_w2p"
                                )
                            os.remove(expected_pdf_path)
                        else:
                            st.error("Falha ao gerar o PDF. Verifique o arquivo enviado.")
                            
                        if os.path.exists(tmp_docx_path):
                            os.remove(tmp_docx_path)
                            
                    except Exception as e:
                        st.error(f"Erro durante a conversão: {str(e)}")
                        
    # OPÇÃO 2: PDF PARA WORD
    else:
        st.subheader("Converter PDF para Word (.docx)")
        pdf_to_convert = st.file_uploader(
            "Selecione um arquivo PDF (.pdf)", 
            type=["pdf"], 
            accept_multiple_files=False,
            key="p2w_uploader"
        )
        
        if pdf_to_convert:
            if st.button("🚀 Converter para Word (.docx)", type="primary", key="btn_p2w"):
                with st.spinner("Convertendo PDF para Word (extraindo textos e layouts)..."):
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                            tmp_pdf.write(pdf_to_convert.getvalue())
                            tmp_pdf_path = tmp_pdf.name
                            
                        output_docx_path = tmp_pdf_path.replace(".pdf", ".docx")
                        
                        cv = Converter(tmp_pdf_path)
                        cv.convert(output_docx_path, start=0, end=None)
                        cv.close()
                        
                        if os.path.exists(output_docx_path):
                            st.success("✅ Conversão para Word (.docx) concluída!")
                            with open(output_docx_path, "rb") as fp:
                                original_base_name, _ = os.path.splitext(pdf_to_convert.name)
                                st.download_button(
                                    label="📥 Baixar Arquivo Word (.docx)",
                                    data=fp.read(),
                                    file_name=f"{original_base_name}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    key="btn_dl_p2w"
                                )
                            os.remove(output_docx_path)
                        else:
                            st.error("Falha ao gerar o arquivo Word.")
                            
                        if os.path.exists(tmp_pdf_path):
                            os.remove(tmp_pdf_path)
                            
                    except Exception as e:
                        st.error(f"Erro durante a conversão de PDF para Word: {str(e)}")

# Rodapé com nota de privacidade
st.markdown("---")
st.caption("🔒 Seus arquivos são processados temporariamente na memória do servidor e descartados imediatamente após o uso. Nenhuma cópia é armazenada.")
