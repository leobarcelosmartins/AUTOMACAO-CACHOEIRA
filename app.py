import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import fitz  # PyMuPDF
import io
import os
import json
import shutil
import subprocess
import tempfile
import pandas as pd
from streamlit_paste_button import paste_image_button
from PIL import Image
import platform
import time
import zipfile
from pathlib import Path

# --- CONFIGURAÇÕES DE LAYOUT ---
st.set_page_config(page_title="Gerador de Relatórios Madalena", layout="wide")

# --- CUSTOM CSS ---
st.markdown("""
    <style>
    .main { background-color: #f0f2f5; }
    
    /* CONFIGURAÇÃO DO GHOST CARD VIA CONTAINER NATIVO */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #ffffff !important;
        border: 1px solid #d1d5db !important;
        border-radius: 12px !important;
        padding: 25px !important;
        margin-bottom: 20px !important;
    }
    
    /* BOTÃO PRIMÁRIO VERDE (CONFORME IMAGEM) */
    div.stButton > button[kind="primary"] {
        background-color: #28a745 !important;
        color: white !important;
        border: none !important;
        width: 100% !important;
        font-weight: bold !important;
        height: 3em !important;
        border-radius: 8px !important;
    }
    
    /* BOTÃO SECUNDÁRIO PADRÃO (BORDA CINZA) */
    div.stButton > button[kind="secondary"] {
        background-color: #ffffff !important;
        border: 1px solid #d1d5db !important;
        height: 3em !important;
        width: 100% !important;
        color: #374151 !important;
    }

    div.stButton > button[key*="del_"] {
        border: 1px solid #dc3545 !important;
        color: #dc3545 !important;
        background-color: transparent !important;
        font-size: 0.8em !important;
        height: 2em !important;
    }
    .upload-label { font-weight: bold; color: #1f2937; margin-bottom: 8px; display: block; }
    </style>
    """, unsafe_allow_html=True)

# --- DICIONÁRIO DE DIMENSÕES (CONTRATO CACHOEIRA) ---
DIMENSOES_CAMPOS = {
    "PRINT_ATEND_OCUPACAO": 165, "PRINT_CLASSIFICAÇÃO": 165,
    "GRAFICO_CIRURGIAS_ELETIVAS": 125, "TABELA_CIRURGIAS": 190,
    "TABELA_RAIOX": 290, "TABELA_CONS_TRANSFERENCIA": 140,
    "TABELA_DET_TRANSFERENCIA": 185, "TABELA_OBITO": 190,
    "ATA_OBITO": 165, "TABELA_CCIH": 100,
    "ATA_COMISSAO_CCIH": 165, "ATA_COMISSAO_REVISAO_PRONT": 165,
    "APERFEICOAMENTO_PROFISSIONAL": 165, "H_TABELA_PESQUISA_INTERNA": 65,
    "H_GRAFICO_PESQUISA_INTERNA": 165, "H_TABELA_PESQUISA_RECEPTIVA": 65,
    "H_GRAFICO_PESQUISA_RECEPTIVA": 165, "H_GRAFICO_PESQUISA_RECEPTIVA_2": 185,
    "UPA_TABELA_ATENDIMENTOS": 180, "UPA_TABELA_CLASSIFICAÇÃO": 160,
    "UPA_RELATORIO_MENSAL_RX": 280, "UPA_TABELA_TRANSFERENCIA": 180,
    "UPA_TABELA_OBITO": 180, "UPA_ATA_OBITO": 170,
    "UPA_ATA_PRONTUARIO": 170, "UPA_ATA_CCIH": 170,
    "UPA_APERF_PROF": 170, "UPA_TABELA_PESQUISA_INTERNA": 65,
    "UPA_GRAFICO_PESQUISA_INTERNA": 140, "UPA_TABELA_PESQUISA_RECEPTIVA": 65,
    "UPA_GRAFICO_PESQUISA_RECEPTIVA": 165, "UPA_GRAFICO_PESQUISA_RECEPTIVA_2": 185,
    "TABELA_QUANTI": 200, "TABELA_QUALI": 200
}

# --- DICIONÁRIO DE DESCRIÇÕES AMIGÁVEIS (LABELS) ---
LABELS_EVIDENCIAS = {
    "PRINT_ATEND_OCUPACAO": "Tabela de Atendimento por Ocupação",
    "PRINT_CLASSIFICAÇÃO": "Tabela de Classificação de Risco",
    "GRAFICO_CIRURGIAS_ELETIVAS": "Gráfico das Cirurgias Eletivas",
    "TABELA_CIRURGIAS": "Tabela de Cirurgias por Profissional",
    "TABELA_RAIOX": "Tabela de Raio X",
    "TABELA_CONS_TRANSFERENCIA": "Tabela Consolidada de Transferências",
    "TABELA_DET_TRANSFERENCIA": "Tabela Detalhada de Transferências",
    "TABELA_OBITO": "Tabela de Óbitos",
    "ATA_OBITO": "Ata de Revisão de Óbito",
    "TABELA_CCIH": "Tabela CCIH",
    "ATA_COMISSAO_CCIH": "Ata Comissão CCIH",
    "ATA_COMISSAO_REVISAO_PRONT": "Ata Comissão Revisão de Prontuário",
    "APERFEICOAMENTO_PROFISSIONAL": "Aperfeiçoamento Profissional (Arquivos)",
    "H_TABELA_PESQUISA_INTERNA": "Tabela da Pesquisa Interna do SAU",
    "H_GRAFICO_PESQUISA_INTERNA": "Gráfico da Pesquisa Interna do SAU",
    "H_TABELA_PESQUISA_RECEPTIVA": "Tabela da Pesquisa Receptiva do SAU",
    "H_GRAFICO_PESQUISA_RECEPTIVA": "Gráfico da Pesquisa Receptiva do SAU",
    "H_GRAFICO_PESQUISA_RECEPTIVA_2": "Tabela 2 da Pesquisa Receptiva do SAU",
    "UPA_TABELA_ATENDIMENTOS": "Tabela de Atendimentos por Ocupação UPA",
    "UPA_TABELA_CLASSIFICAÇÃO": "Tabela de Classificação de Risco UPA",
    "UPA_RELATORIO_MENSAL_RX": "Relatório Mensal Raio X UPA",
    "UPA_TABELA_TRANSFERENCIA": "Tabela de Transferência UPA",
    "UPA_TABELA_OBITO": "Tabela de Óbitos UPA",
    "UPA_ATA_OBITO": "Ata de Revisão de Óbito UPA",
    "UPA_ATA_PRONTUARIO": "Ata de Revisão de Prontuário UPA",
    "UPA_ATA_CCIH": "Ata de Revisão da Comissão do CCIH UPA",
    "UPA_APERF_PROF": "UPA Aperfeiçoamento Profissional",
    "UPA_TABELA_PESQUISA_INTERNA": "Tabela da Pesquisa Interna UPA",
    "UPA_GRAFICO_PESQUISA_INTERNA": "Gráfico da Pesquisa Interna UPA",
    "UPA_TABELA_PESQUISA_RECEPTIVA": "Tabela da Pesquisa Receptiva UPA",
    "UPA_GRAFICO_PESQUISA_RECEPTIVA": "Gráfico da Pesquisa Receptiva UPA",
    "UPA_GRAFICO_PESQUISA_RECEPTIVA_2": "Tabela 2 da Pesquisa Receptiva UPA",
    "TABELA_QUANTI": "Tabela Quanti (Geral)",
    "TABELA_QUALI": "Tabela Quali (Geral)"
}

# --- CHAVES DE CAMPOS ---
FORM_KEYS = [
    "sel_mes", "sel_ano", "H_T_PAC_INT", "H_ALTA", "H_TRANSF_MAIOR", "H_TRANSF_MENOR",
    "H_TRANSF_INT", "H_EVASAO", "H_OBITO_MAIOR", "H_OBITO_MENOR", "H_OB_INT",
    "H_GINECO", "H_CIR_GERAL", "H_MED_CLI", "H_ORTO", "H_PED",
    "AMB_FISIO", "AMB_PSICO", "AMB_FONO", "AMB_SERV_SOC",
    "PARECER_ASSIST_SOC", "PARECER_NEURO", "PARECER_CLI_GER", "PARECER_CIR_GER", "PARECER_URO",
    "PARECER_NEURO_CIR", "PARECER_PSICO", "PARECER_CARDIO", "PARECER_HEMA", "PARECER_NEFRO",
    "PARECER_PSIQ", "PARECER_CIR_VASC", "PARECER_BUCO", "PARECER_OBSTR", "PARECER_OTORRINO", "PARECER_DERMA",
    "H_ELE_CIR_GER", "H_ELE_CIR_ORTO", "H_ELE_CIR_BUCO", "H_ELE_CIR_URO",
    "H_EMERG_CIR_GER", "H_EMERG_PART_CES", "H_EMERG_VASC", "H_EMERG_URO", "H_EMERG_ORT", "H_EMERG_GINECO",
    "H_PF_LAQ", "H_PF_DIU", "H_PF_BIO", "H_EX_ENDO", "H_EX_COLO",
    "AMB_EX_HEMOD", "AMB_EX_LABOR", "AMB_EX_RADIO", "H_RP_TOTAL_PAC", "H_SAU_PESQ_INT", "H_SAU_OUV_RECEP",
    "UPA_MED_CLI", "UPA_MED_PED", "UPA_ATEND_AS", "UPA_ATEND_NUTRI", "UPA_EX_ELETRO", "UPA_EX_LAB", "UPA_EX_RADIO",
    "UPA_PESQ_INT", "UPA_PESQ_RECEP", "UPA_T_TRANSF"
]

# --- ESTADO DA SESSÃO ---
if 'dados_sessao' not in st.session_state:
    st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}
if 'relatorio_atual' not in st.session_state:
    st.session_state.relatorio_atual = ""

BASE_RELATORIOS_DIR = Path("relatorios_guardados")
BASE_RELATORIOS_DIR.mkdir(exist_ok=True)

# --- FUNÇÕES DE PERSISTÊNCIA (DISCO LOCAL) ---
def _normalizar_nome(nome):
    return "".join([c if c.isalnum() else "_" for c in nome.strip()])

def salvar_relatorio(nome):
    if not nome:
        st.warning("Informe um nome para o relatório.")
        return
    pasta = BASE_RELATORIOS_DIR / _normalizar_nome(nome)
    pasta.mkdir(exist_ok=True)
    evid_meta = {}
    pasta_evid = pasta / "evidencias"
    pasta_evid.mkdir(exist_ok=True)
    for m, itens in st.session_state.dados_sessao.items():
        evid_meta[m] = []
        for i, item in enumerate(itens):
            _, ext = os.path.splitext(item["name"])
            if not ext: ext = ".png"
            nome_arq = f"{m}_{i}{ext}"
            conteudo = item["content"]
            if isinstance(conteudo, Image.Image):
                conteudo.save(pasta_evid / nome_arq, format="PNG")
            else:
                if hasattr(conteudo, "getvalue"): data = conteudo.getvalue()
                else: 
                    conteudo.seek(0)
                    data = conteudo.read()
                with open(pasta_evid / nome_arq, "wb") as f: f.write(data)
            evid_meta[m].append({"name": item["name"], "file": f"evidencias/{nome_arq}", "type": item["type"]})
    estado = {"form_state": {k: st.session_state.get(k) for k in FORM_KEYS}, "evidencias": evid_meta}
    with open(pasta / "estado.json", "w", encoding="utf-8") as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)
    st.session_state.relatorio_atual = nome
    st.success(f"Relatório '{nome}' guardado em disco!")

def carregar_relatorio(nome):
    pasta = BASE_RELATORIOS_DIR / nome
    with open(pasta / "estado.json", "r", encoding="utf-8") as f:
        estado = json.load(f)
    for k, v in estado["form_state"].items():
        st.session_state[k] = v
    st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}
    for m, lista in estado["evidencias"].items():
        for meta in lista:
            with open(pasta / meta["file"], "rb") as f: data = f.read()
            bio = io.BytesIO(data)
            bio.name = meta["name"]
            st.session_state.dados_sessao[m].append({"name": meta["name"], "content": bio, "type": meta["type"]})
    st.session_state.relatorio_atual = nome
    st.toast(f"Relatório '{nome}' carregado.")

def excluir_relatorio(nome):
    pasta = BASE_RELATORIOS_DIR / nome
    if pasta.exists():
        shutil.rmtree(pasta)
        st.success(f"Relatório '{nome}' excluído.")
        st.rerun()

# --- FUNÇÕES DE BACKUP (ZIP) ---
def gerar_backup_zip():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        evid_meta = {}
        for marcador, itens in st.session_state.dados_sessao.items():
            evid_meta[marcador] = []
            for i, item in enumerate(itens):
                conteudo = item["content"]
                _, ext = os.path.splitext(item["name"])
                if not ext: ext = ".png"
                file_bytes = b""
                if isinstance(conteudo, Image.Image):
                    img_buf = io.BytesIO()
                    conteudo.save(img_buf, format="PNG")
                    file_bytes = img_buf.getvalue()
                else:
                    if hasattr(conteudo, "getvalue"): file_bytes = conteudo.getvalue()
                    elif hasattr(conteudo, "read"): 
                        conteudo.seek(0)
                        file_bytes = conteudo.read()
                    else: file_bytes = conteudo
                nome_interno = f"evidencias/{marcador}_{i}{ext}"
                zf.writestr(nome_interno, file_bytes)
                evid_meta[marcador].append({"name": item["name"], "file": nome_interno, "type": item["type"]})
        estado = {"form_state": {k: st.session_state.get(k) for k in FORM_KEYS}, "evidencias": evid_meta}
        zf.writestr("estado.json", json.dumps(estado, ensure_ascii=False, indent=2))
    buf.seek(0)
    return buf

def processar_upload_backup(uploaded_zip):
    try:
        with zipfile.ZipFile(uploaded_zip, "r") as zf:
            estado_str = zf.read("estado.json").decode("utf-8")
            estado = json.loads(estado_str)
            for k, v in estado.get("form_state", {}).items(): st.session_state[k] = v
            st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}
            for marcador, lista in estado.get("evidencias", {}).items():
                for meta in lista:
                    try:
                        file_bytes = zf.read(meta["file"])
                        bio = io.BytesIO(file_bytes)
                        bio.name = meta["name"]
                        st.session_state.dados_sessao[marcador].append({"name": meta["name"], "content": bio, "type": meta["type"]})
                    except: pass
        st.success("✅ Backup restaurado com sucesso!")
    except Exception as e:
        st.error(f"Erro ao ler backup: {e}")

# --- FUNÇÕES CORE ---
def converter_para_pdf(docx_path, output_dir):
    comando = 'libreoffice'
    if platform.system() == "Windows":
        caminhos = ['libreoffice', r'C:\Program Files\LibreOffice\program\soffice.exe', r'C:\Program Files (x86)\LibreOffice\program\soffice.exe']
        for p in caminhos:
            try:
                subprocess.run([p, '--version'], capture_output=True, check=True)
                comando = p; break
            except: continue
    subprocess.run([comando, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, docx_path], check=True)

def processar_item_lista(doc_template, item, marcador):
    largura = DIMENSOES_CAMPOS.get(marcador, 165)
    try:
        if isinstance(item, Image.Image):
            img_buf = io.BytesIO()
            item.save(img_buf, format='PNG')
            img_buf.seek(0)
            return [InlineImage(doc_template, img_buf, width=Mm(largura))]
        if hasattr(item, 'seek'): item.seek(0)
        ext = getattr(item, 'name', '').lower()
        if ext.endswith(".pdf"):
            pdf = fitz.open(stream=item.read(), filetype="pdf")
            imgs = []
            for pg in pdf:
                pix = pg.get_pixmap(matrix=fitz.Matrix(2, 2))
                imgs.append(InlineImage(doc_template, io.BytesIO(pix.tobytes()), width=Mm(largura)))
            pdf.close(); return imgs
        return [InlineImage(doc_template, item, width=Mm(largura))]
    except: return []

# --- UI PRINCIPAL ---
st.title("Automação de Relatórios - Cachoeira")

# --- GESTOR DE RELATÓRIOS (ESTILO DA IMAGEM) ---
with st.expander("📂 Gestor de Relatórios Guardados", expanded=not st.session_state.relatorio_atual):
    col_g1, col_g2 = st.columns([2, 1])
    with col_g1:
        lista_pastas = [p.name for p in BASE_RELATORIOS_DIR.iterdir() if p.is_dir()]
        sel_disco = st.selectbox("Relatórios Guardados", ["-- Selecionar --"] + lista_pastas)
        ca1, ca2 = st.columns(2)
        if ca1.button("📥 Carregar Selecionado", use_container_width=True) and sel_disco != "-- Selecionar --":
            carregar_relatorio(sel_disco)
            st.rerun()
        if ca2.button("🗑️ Excluir Selecionado", use_container_width=True) and sel_disco != "-- Selecionar --":
            excluir_relatorio(sel_disco)
    with col_g2:
        novo_nome = st.text_input("Nome do Relatório", placeholder="Ex: Pacheco_Marco_2025", value=st.session_state.relatorio_atual)
        if st.button("💾 Salvar Progresso", use_container_width=True, type="primary"):
            salvar_relatorio(novo_nome)

# --- BACKUP EXTERNO (ZIP) - MESMO ESTILO ---
with st.expander("☁️ Backup Externo (Importar / Exportar ZIP)", expanded=False):
    col_z1, col_z2 = st.columns([2, 1])
    with col_z1:
        zip_upload = st.file_uploader("Relatório em Arquivo ZIP", type=["zip"], key="zip_up", label_visibility="collapsed")
        if st.button("📥 Restaurar do Arquivo ZIP", use_container_width=True) and zip_upload:
            processar_upload_backup(zip_upload)
            time.sleep(1)
            st.rerun()
    with col_z2:
        st.markdown("<div style='height: 2px;'></div>", unsafe_allow_html=True) # Alinhamento visual
        zip_data = gerar_backup_zip()
        st.download_button(
            label="📤 Baixar Backup ZIP",
            data=zip_data,
            file_name=f"Backup_Cachoeira_{st.session_state.get('sel_mes', 'Atual')}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

t_hosp, t_amb, t_cir, t_upa, t_evidencia = st.tabs(
    ["HOSPITAL", "AMBULATÓRIO/PARECER", "CIRURGIAS/EXAMES", "UPA", "ARQUIVOS"]
)

# --- ABA HOSPITAL ---
with t_hosp:
    with st.container(border=True):
        st.markdown("### Período e Internação")
        c1, c2, c3 = st.columns(3)
        with c1: st.selectbox("Mês", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"], key="sel_mes")
        with c2: st.selectbox("Ano", [2024, 2025, 2026], index=1, key="sel_ano")
        with c3: st.number_input("Total Pacientes Internados", key="H_T_PAC_INT", step=1)
    
    with st.container(border=True):
        st.markdown("### Saídas e Óbitos")
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.number_input("Altas", key="H_ALTA", step=1)
        with c2: st.number_input("Transf > 24H", key="H_TRANSF_MAIOR", step=1)
        with c3: st.number_input("Transf < 24H", key="H_TRANSF_MENOR", step=1)
        with c4: st.number_input("Evasão", key="H_EVASAO", step=1)
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.number_input("Óbito > 24H", key="H_OBITO_MAIOR", step=1)
        with c2: st.number_input("Óbito < 24H", key="H_OBITO_MENOR", step=1)
        with c3: st.number_input("Transf Internação", key="H_TRANSF_INT", step=1)
        with c4: st.number_input("Óbito Internação", key="H_OB_INT", step=1)

    with st.container(border=True):
        st.markdown("### Emergência Hospitalar")
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: st.number_input("Ginecologia", key="H_GINECO", step=1)
        with c2: st.number_input("Cirurgia Geral", key="H_CIR_GERAL", step=1)
        with c3: st.number_input("Médico Clínico", key="H_MED_CLI", step=1)
        with c4: st.number_input("Ortopedia", key="H_ORTO", step=1)
        with c5: st.number_input("Pediatria", key="H_PED", step=1)

# --- ABA AMBULATÓRIO E PARECERES ---
with t_amb:
    with st.container(border=True):
        st.markdown("### Atendimentos Ambulatoriais")
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.number_input("Fisioterapia", key="AMB_FISIO", step=1)
        with c2: st.number_input("Psicologia", key="AMB_PSICO", step=1)
        with c3: st.number_input("Fonoaudiologia", key="AMB_FONO", step=1)
        with c4: st.number_input("Serviço Social", key="AMB_SERV_SOC", step=1)
    with st.container(border=True):
        st.markdown("### Pareceres Médicos")
        cp1, cp2, cp3, cp4 = st.columns(4)
        with cp1:
            st.number_input("Assist. Social", key="PARECER_ASSIST_SOC", step=1)
            st.number_input("Neurologia", key="PARECER_NEURO", step=1)
            st.number_input("Clínico Geral", key="PARECER_CLI_GER", step=1)
            st.number_input("Cirurgia Geral", key="PARECER_CIR_GER", step=1)
        with cp2:
            st.number_input("Urologia", key="PARECER_URO", step=1)
            st.number_input("Neurocirurgia", key="PARECER_NEURO_CIR", step=1)
            st.number_input("Psicólogo", key="PARECER_PSICO", step=1)
            st.number_input("Cardiologista", key="PARECER_CARDIO", step=1)
        with cp3:
            st.number_input("Hematologia", key="PARECER_HEMA", step=1)
            st.number_input("Nefrologia", key="PARECER_NEFRO", step=1)
            st.number_input("Psiquiatria", key="PARECER_PSIQ", step=1)
            st.number_input("Cir. Vascular", key="PARECER_CIR_VASC", step=1)
        with cp4:
            st.number_input("Bucomaxilo", key="PARECER_BUCO", step=1)
            st.number_input("Obstetra", key="PARECER_OBSTR", step=1)
            st.number_input("Otorrino", key="PARECER_OTORRINO", step=1)
            st.number_input("Dermatologia", key="PARECER_DERMA", step=1)

# --- ABA CIRURGIAS E EXAMES ---
with t_cir:
    with st.container(border=True):
        st.markdown("### Cirurgias Eletivas")
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.number_input("Cirurgia Geral", key="H_ELE_CIR_GER", step=1)
        with c2: st.number_input("Ortopedia", key="H_ELE_CIR_ORTO", step=1)
        with c3: st.number_input("Bucomaxilo", key="H_ELE_CIR_BUCO", step=1)
        with c4: st.number_input("Urologia", key="H_ELE_CIR_URO", step=1)
    with st.container(border=True):
        st.markdown("### Cirurgias de Emergência")
        c1, c2, c3 = st.columns(3)
        with c1: st.number_input("Emerg. Cirurgia Geral", key="H_EMERG_CIR_GER", step=1)
        with c2: st.number_input("Emerg. Parto Cesárea", key="H_EMERG_PART_CES", step=1)
        with c3: st.number_input("Emerg. Vascular", key="H_EMERG_VASC", step=1)
    with st.container(border=True):
        st.markdown("### Planejamento Familiar e Exames")
        c1, c2, c3 = st.columns(3)
        with c1: st.number_input("Laqueadura", key="H_PF_LAQ", step=1)
        with c2: st.number_input("Retirada de DIU", key="H_PF_DIU", step=1)
        with c3: st.number_input("Biópsia", key="H_PF_BIO", step=1)
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: st.number_input("Endoscopia", key="H_EX_ENDO", step=1)
        with c2: st.number_input("Colonoscopia", key="H_EX_COLO", step=1)
        with c3: st.number_input("Hemodiálise", key="AMB_EX_HEMOD", step=1)
        with c4: st.number_input("Laboratório", key="AMB_EX_LABOR", step=1)
        with c5: st.number_input("Radiografia", key="AMB_EX_RADIO", step=1)
    with st.container(border=True):
        st.markdown("### Pesquisa SAU e Revisão")
        c1, c2, c3 = st.columns(3)
        with c1: st.number_input("Revisão de Prontuário", key="H_RP_TOTAL_PAC", step=1)
        with c2: st.number_input("Pesquisa Interna", key="H_SAU_PESQ_INT", step=1)
        with c3: st.number_input("Ouvidoria Receptiva", key="H_SAU_OUV_RECEP", step=1)

# --- ABA UPA ---
with t_upa:
    with st.container(border=True):
        st.markdown("### Atendimentos UPA")
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.number_input("Médico Clínico UPA", key="UPA_MED_CLI", step=1)
        with c2: st.number_input("Médico Pediatra UPA", key="UPA_MED_PED", step=1)
        with c3: st.number_input("Assistente Social UPA", key="UPA_ATEND_AS", step=1)
        with c4: st.number_input("Nutricionista UPA", key="UPA_ATEND_NUTRI", step=1)
    with st.container(border=True):
        st.markdown("### Exames e Transferências UPA")
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.number_input("Eletrocardiograma UPA", key="UPA_EX_ELETRO", step=1)
        with c2: st.number_input("Laboratório UPA", key="UPA_EX_LAB", step=1)
        with c3: st.number_input("Radiografia UPA", key="UPA_EX_RADIO", step=1)
        with c4: st.number_input("Total Transferências UPA", key="UPA_T_TRANSF", step=1)

# --- ABA ARQUIVOS (EVIDÊNCIAS) ---
with t_evidencia:
    # Agrupamento para facilitar a navegação (Opcional, mas recomendado)
    secoes_evidencias = [
        {"nome": "Hospital - Atendimentos e Classificação", "marcadores": ["PRINT_ATEND_OCUPACAO", "PRINT_CLASSIFICAÇÃO"]},
        {"nome": "Hospital - Cirurgias e Procedimentos", "marcadores": ["GRAFICO_CIRURGIAS_ELETIVAS", "TABELA_CIRURGIAS", "TABELA_RAIOX", "H_T_PROC_CIR"]},
        {"nome": "Hospital - Transferências e Óbitos", "marcadores": ["TABELA_CONS_TRANSFERENCIA", "TABELA_DET_TRANSFERENCIA", "TABELA_OBITO", "ATA_OBITO"]},
        {"nome": "Hospital - Comissões e Qualidade", "marcadores": ["TABELA_CCIH", "ATA_COMISSAO_CCIH", "ATA_COMISSAO_REVISAO_PRONT", "APERFEICOAMENTO_PROFISSIONAL"]},
        {"nome": "Hospital - Pesquisa de Satisfação (SAU)", "marcadores": ["H_TABELA_PESQUISA_INTERNA", "H_GRAFICO_PESQUISA_INTERNA", "H_TABELA_PESQUISA_RECEPTIVA", "H_GRAFICO_PESQUISA_RECEPTIVA", "H_GRAFICO_PESQUISA_RECEPTIVA_2"]},
        {"nome": "UPA - Atendimentos e Exames", "marcadores": ["UPA_TABELA_ATENDIMENTOS", "UPA_TABELA_CLASSIFICAÇÃO", "UPA_RELATORIO_MENSAL_RX", "UPA_TABELA_TRANSFERENCIA", "UPA_TABELA_OBITO"]},
        {"nome": "UPA - Comissões e Qualidade", "marcadores": ["UPA_ATA_OBITO", "UPA_ATA_PRONTUARIO", "UPA_ATA_CCIH", "UPA_APERF_PROF"]},
        {"nome": "UPA - Pesquisa de Satisfação", "marcadores": ["UPA_TABELA_PESQUISA_INTERNA", "UPA_GRAFICO_PESQUISA_INTERNA", "UPA_TABELA_PESQUISA_RECEPTIVA", "UPA_GRAFICO_PESQUISA_RECEPTIVA", "UPA_GRAFICO_PESQUISA_RECEPTIVA_2"]},
        {"nome": "Indicadores Gerais", "marcadores": ["TABELA_QUANTI", "TABELA_QUALI"]}
    ]

    for secao in secoes_evidencias:
        with st.expander(f"📌 {secao['nome']}", expanded=False):
            for marcador in secao['marcadores']:
                # Verificação caso o marcador exista no dicionário de dimensões
                if marcador in DIMENSOES_CAMPOS:
                    with st.container(border=True):
                        # Uso do dicionário LABELS_EVIDENCIAS para mostrar o nome amigável
                        label_exibicao = LABELS_EVIDENCIAS.get(marcador, marcador)
                        st.markdown(f"<span class='upload-label'>{label_exibicao} (Largura: {DIMENSOES_CAMPOS[marcador]}mm)</span>", unsafe_allow_html=True)
                        
                        f_up = st.file_uploader("Upload", type=['png', 'jpg', 'pdf'], key=f"f_{marcador}", label_visibility="collapsed")
                        if f_up and f_up.name not in [x['name'] for x in st.session_state.dados_sessao.get(marcador, [])]:
                            st.session_state.dados_sessao[marcador].append({"name": f_up.name, "content": f_up, "type": "f"})
                        
                        kp = f"p_{marcador}_{len(st.session_state.dados_sessao.get(marcador, []))}"
                        pasted = paste_image_button(label="📸 Colar Print", key=kp)
                        if pasted is not None and pasted.image_data is not None:
                            st.session_state.dados_sessao[marcador].append({"name": f"Captura_{marcador}_{int(time.time())}.png", "content": pasted.image_data, "type": "p"})
                            st.toast(f"Anexado: {label_exibicao}"); time.sleep(0.4); st.rerun()
                        
                        if st.session_state.dados_sessao.get(marcador):
                            for idx, item in enumerate(st.session_state.dados_sessao[marcador]):
                                col1, col2 = st.columns([0.9, 0.1])
                                col1.caption(f"📄 {item['name']}")
                                if col2.button("🗑️", key=f"del_{marcador}_{idx}"): st.session_state.dados_sessao[marcador].pop(idx); st.rerun()

# --- GERAÇÃO FINAL ---
if st.button("FINALIZAR E GERAR RELATÓRIO CACHOEIRA", type="primary", key="btn_finalizar"):
    try:
        with st.spinner("Gerando documento..."):
            with tempfile.TemporaryDirectory() as tmp:
                doc = DocxTemplate("template-cachoeira.docx")
                # Lógica de somas para Cachoeira
                h_total_saida = sum([st.session_state.get(k, 0) for k in ["H_ALTA", "H_TRANSF_MAIOR", "H_TRANSF_MENOR", "H_EVASAO", "H_OBITO_MAIOR", "H_OBITO_MENOR"]])
                h_t_atend_emerg = sum([st.session_state.get(k, 0) for k in ["H_GINECO", "H_CIR_GERAL", "H_MED_CLI", "H_ORTO", "H_PED"]])
                total_amb_parecer = sum([st.session_state.get(k, 0) for k in FORM_KEYS if "PARECER_" in k])
                contexto = {k: st.session_state.get(k, 0) for k in FORM_KEYS}
                contexto.update({"SISTEMA_MES_REFERENCIA": f"{st.session_state.sel_mes}/{st.session_state.sel_ano}", "H_TOTAL_SAIDA": h_total_saida, "H_T_ATEND_EMERG": h_t_atend_emerg, "TOTAL_AMB_PARECER": total_amb_parecer})
                for m in DIMENSOES_CAMPOS.keys():
                    imgs = []
                    for item in st.session_state.dados_sessao.get(m, []):
                        res = processar_item_lista(doc, item['content'], m)
                        if res: imgs.extend(res)
                    contexto[m] = imgs
                doc.render(contexto); docx_p = os.path.join(tmp, "relatorio.docx"); doc.save(docx_p)
                st.success("✅ Relatório gerado!"); cd1, cd2 = st.columns(2)
                with cd1: st.download_button("WORD (.docx)", open(docx_p, "rb").read(), f"RELATORIO_CACHOEIRA_{st.session_state.sel_mes}.docx")
                with cd2: 
                    try: 
                        converter_para_pdf(docx_p, tmp)
                        st.download_button("PDF", open(os.path.join(tmp, "relatorio.pdf"), "rb").read(), f"RELATORIO_CACHOEIRA_{st.session_state.sel_mes}.pdf")
                    except: st.warning("PDF falhou.")
    except Exception as e: st.error(f"Erro: {e}")

st.caption("Desenvolvido por Leonardo Barcelos Martins")
