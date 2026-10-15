# Mesclador de Documentos Word e PDF (Streamlit)

Aplicacao web desenvolvida com Streamlit para unir multiplos arquivos Word (.docx) e arquivos PDF (.pdf) em um unico documento final.

## Como Publicar no Streamlit Cloud (100% Gratis)

Voce nao precisa instalar nada no seu computador! Siga os passos abaixo:

### Passo 1: Criar Repositorio no GitHub
1. Acesse o GitHub e faca login na sua conta.
2. Clique em New Repository (Novo Repositorio).
3. Defina um nome (ex: mesclador-documentos-streamlit) e escolha Public.
4. Clique em Create repository.
5. Faca o upload dos arquivos desta pasta (app.py, requirements.txt, packages.txt) para o repositorio.

### Passo 2: Conectar ao Streamlit Community Cloud
1. Acesse o Streamlit Community Cloud (share.streamlit.io).
2. Faca login usando sua conta do GitHub.
3. Clique no botao New app.
4. Selecione o repositorio que voce acabou de criar no GitHub.
5. No campo Main file path, certifique-se de que esta escrito app.py.
6. Clique em Deploy!

O Streamlit Cloud vai ler o arquivo packages.txt, instalar o LibreOffice e todas as bibliotecas do Python na nuvem e colocar o seu site no ar em aproximadamente 2 minutos!

---

## Funcionalidades
- Mesclagem de Word (.docx): Une 30+ documentos Word preservando imagens, tabelas e estilos via docxcompose.
- Conversao para PDF: Gera a versao PDF do Word unificado via LibreOffice Headless.
- Mesclagem de PDFs (.pdf): Junta 50+ PDFs soltos em um unico arquivo PDF via pypdf.
- Reordenacao: Permite alterar a ordem sequencial dos documentos antes da fusao.
