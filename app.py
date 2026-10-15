import os
import tempfile
import subprocess
import streamlit as st
from docx import Document
from docxcompose.composer import Composer
from pypdf import PdfMerger

st.set_page_config(
    page_title="Mesclador de Documentos Word & PDF",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Mesclador de Documentos Word & PDF")
st.markdown("Junte múltiplos documentos Word (.docx) ou arquivos PDF em um único arquivo de forma simples e rápida.")

tab1, tab2 = st.tabs(["📝 Mesclar Arquivos Word (.docx)", "📕 Mesclar Arquivos PDF (.pdf)"])

with tab1:
    st.header("Mesclar Documentos Word")
    st.write("Envie seus arquivos .docx (ex: 30 ou mais arquivos). Eles serão unificados mantendo a formatação original.")
    
    uploaded_docx_files = st.file_uploader(
        "Selecione os arquivos Word (.docx)",
        type=["docx"],
        accept_multiple_files=True,
        key="docx_uploader"
    )
    
    if uploaded_docx_files:
        st.success(f"{len(uploaded_docx_files)} arquivo(s) selecionado(s).")
        
        st.subheader("Ordem de Junção")
        st.caption("Ajuste a ordem dos arquivos se necessário:")
        
        filenames = [f.name for f in uploaded_docx_files]
        file_map = {f.name: f for f in uploaded_docx_files}
        
        if "docx_order" not in st.session_state or set(st.session_state.docx_order) != set(filenames):
            st.session_state.docx_order = filenames
            
        for i, name in enumerate(st.session_state.docx_order):
            col1, col2, col3 = st.columns([6, 1, 1])
            col1.write(f"**{i+1}.** {name}")
            if col2.button("⬆️", key=f"up_docx_{i}") and i > 0:
                st.session_state.docx_order[i], st.session_state.docx_order[i-1] = st.session_state.docx_order[i-1], st.session_state.docx_order[i]
                st.rerun()
            if col3.button("⬇️", key=f"down_docx_{i}") and i < len(st.session_state.docx_order) - 1:
                st.session_state.docx_order[i], st.session_state.docx_order[i+1] = st.session_state.docx_order[i+1], st.session_state.docx_order[i]
                st.rerun()

        st.divider()
        out_format = st.radio("Formato de Saída:", ["Apenas Word (.docx)", "Apenas PDF (.pdf)", "Ambos (.docx e .pdf)"], index=2)
        
        if st.button("🚀 Processar e Juntar Arquivos Word", type="primary"):
            with st.spinner("Processando e unificando documentos Word..."):
                try:
                    with tempfile.TemporaryDirectory() as temp_dir:
                        ordered_files = [file_map[name] for name in st.session_state.docx_order]
                        
                        saved_paths = []
                        for idx, uf in enumerate(ordered_files):
                            p = os.path.join(temp_dir, f"doc_{idx}.docx")
                            with open(p, "wb") as f:
                                f.write(uf.getvalue())
                            saved_paths.append(p)
                        
                        master_doc = Document(saved_paths[0])
                        composer = Composer(master_doc)
                        
                        for p in saved_paths[1:]:
                            sub_doc = Document(p)
                            composer.append(sub_doc)
                        
                        merged_docx_path = os.path.join(temp_dir, "documento_final.docx")
                        composer.save(merged_docx_path)
                        
                        merged_docx_bytes = None
                        merged_pdf_bytes = None
                        
                        if out_format in ["Apenas Word (.docx)", "Ambos (.docx e .pdf)"]:
                            with open(merged_docx_path, "rb") as f:
                                merged_docx_bytes = f.read()
                        
                        if out_format in ["Apenas PDF (.pdf)", "Ambos (.docx e .pdf)"]:
                            cmd = ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", temp_dir, merged_docx_path]
                            subprocess.run(cmd, check=True)
                            merged_pdf_path = os.path.join(temp_dir, "documento_final.pdf")
                            if os.path.exists(merged_pdf_path):
                                with open(merged_pdf_path, "rb") as f:
                                    merged_pdf_bytes = f.read()
                        
                        st.success("✅ Unificação concluída com sucesso!")
                        
                        c1, c2 = st.columns(2)
                        if merged_docx_bytes:
                            c1.download_button("📥 Baixar Word Final (.docx)", data=merged_docx_bytes, file_name="documento_mesclado.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                        if merged_pdf_bytes:
                            c2.download_button("📥 Baixar PDF Final (.pdf)", data=merged_pdf_bytes, file_name="documento_mesclado.pdf", mime="application/pdf")
                except Exception as e:
                    st.error(f"Erro ao processar os arquivos: {e}")

with tab2:
    st.header("Mesclar Arquivos PDF")
    st.write("Envie seus arquivos PDF (ex: 50 ou mais arquivos soltos). Eles serão unificados em um novo PDF único.")
    
    uploaded_pdf_files = st.file_uploader(
        "Selecione os arquivos PDF (.pdf)",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_uploader"
    )
    
    if uploaded_pdf_files:
        st.success(f"{len(uploaded_pdf_files)} arquivo(s) PDF selecionado(s).")
        
        st.subheader("Ordem de Junção")
        st.caption("Ajuste a ordem dos arquivos se necessário:")
        
        pdf_filenames = [f.name for f in uploaded_pdf_files]
        pdf_file_map = {f.name: f for f in uploaded_pdf_files}
        
        if "pdf_order" not in st.session_state or set(st.session_state.pdf_order) != set(pdf_filenames):
            st.session_state.pdf_order = pdf_filenames
            
        for i, name in enumerate(st.session_state.pdf_order):
            col1, col2, col3 = st.columns([6, 1, 1])
            col1.write(f"**{i+1}.** {name}")
            if col2.button("⬆️", key=f"up_pdf_{i}") and i > 0:
                st.session_state.pdf_order[i], st.session_state.pdf_order[i-1] = st.session_state.pdf_order[i-1], st.session_state.pdf_order[i]
                st.rerun()
            if col3.button("⬇️", key=f"down_pdf_{i}") and i < len(st.session_state.pdf_order) - 1:
                st.session_state.pdf_order[i], st.session_state.pdf_order[i+1] = st.session_state.pdf_order[i+1], st.session_state.pdf_order[i]
                st.rerun()

        st.divider()
        if st.button("🚀 Processar e Juntar PDFs", type="primary"):
            with st.spinner("Unificando arquivos PDF..."):
                try:
                    merger = PdfMerger()
                    for name in st.session_state.pdf_order:
                        f_obj = pdf_file_map[name]
                        merger.append(f_obj)
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_out:
                        merger.write(tmp_out.name)
                        merger.close()
                        tmp_out_path = tmp_out.name
                    
                    with open(tmp_out_path, "rb") as f:
                        final_pdf_bytes = f.read()
                    
                    os.remove(tmp_out_path)
                    
                    st.success("✅ PDFs unificados com sucesso!")
                    st.download_button("📥 Baixar PDF Final Unificado", data=final_pdf_bytes, file_name="pdfs_mesclados.pdf", mime="application/pdf")
                except Exception as e:
                    st.error(f"Erro ao juntar PDFs: {e}")
