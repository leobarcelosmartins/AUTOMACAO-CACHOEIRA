import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import fitz  # PyMuPDF
import io
import os
import subprocess
import tempfile
import pandas as pd
from streamlit_paste_button import paste_image_button
from PIL import Image
import platform
import time
import json
from pathlib import Path

# --- CONFIGURAÇÕES DE LAYOUT ---
st.set_page_config(page_title="Gerador de Relatórios Cachoeira", layout="centered")

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
    
    div.stButton > button[kind="primary"] {
        background-color: #2c86b0 !important;
        color: white !important;
        border: none !important;
        width: 100% !important;
        font-weight: bold !important;
        height: 3em !important;
        border-radius: 8px !important;
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

# --- DICIONÁRIO DE DIMENSÕES DAS EVIDÊNCIAS (CONFORME PDF CACHOEIRA) ---
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

# --- DIRETÓRIO DE RELATÓRIOS SALVOS ---
BASE_RELATORIOS_DIR = Path("relatorios_cachoeira")
BASE_RELATORIOS_DIR.mkdir(exist_ok=True)

# --- CHAVES DE CAMPOS QUE SERÃO PERSISTIDAS (CONTRATO CACHOEIRA) ---
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

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3208/3208726.png", width=100)
    st.title("Painel de Controle")
    st.markdown("---")
    total_anexos = sum(len(v) for v in st.session_state.dados_sessao.values())
    st.metric("Total de Evidências", total_anexos)
    if st.button(" 🗑️ Limpar Todos os Dados", key="btn_limpar_tudo"):
        st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}
        st.rerun()

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
st.caption("Versão 0.9.3")

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
        c1, c2, c3 = st.columns(3)
        with c1: st.number_input("Emerg. Urologia", key="H_EMERG_URO", step=1)
        with c2: st.number_input("Emerg. Ortopedia", key="H_EMERG_ORT", step=1)
        with c3: st.number_input("Emerg. Ginecologia", key="H_EMERG_GINECO", step=1)

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

    with st.container(border=True):
        st.markdown("### Pesquisa de Satisfação UPA")
        c1, c2 = st.columns(2)
        with c1: st.number_input("Pesquisa Interna UPA", key="UPA_PESQ_INT", step=1)
        with c2: st.number_input("Pesquisa Receptiva UPA", key="UPA_PESQ_RECEP", step=1)

# --- ABA ARQUIVOS (EVIDÊNCIAS) ---
with t_evidencia:
    for marcador in DIMENSOES_CAMPOS.keys():
        with st.container(border=True):
            st.markdown(f"<span class='upload-label'>{marcador} (Largura: {DIMENSOES_CAMPOS[marcador]}mm)</span>", unsafe_allow_html=True)
            f_up = st.file_uploader("Upload", type=['png', 'jpg', 'pdf'], key=f"f_{marcador}", label_visibility="collapsed")
            if f_up and f_up.name not in [x['name'] for x in st.session_state.dados_sessao.get(marcador, [])]:
                st.session_state.dados_sessao[marcador].append({"name": f_up.name, "content": f_up, "type": "f"})
            
            # CORREÇÃO DO LOOP INFINITO: Chave dinâmica baseada no número de itens
            kp = f"p_{marcador}_{len(st.session_state.dados_sessao.get(marcador, []))}"
            pasted = paste_image_button(label="📸 Colar Print", key=kp)
            
            if pasted is not None and pasted.image_data is not None:
                st.session_state.dados_sessao[marcador].append({
                    "name": f"Captura_{marcador}_{int(time.time())}.png", 
                    "content": pasted.image_data, 
                    "type": "p"
                })
                # Feedback de sucesso
                st.toast(f"Anexado: {marcador}")
                time.sleep(0.4)
                st.rerun()
            
            if st.session_state.dados_sessao.get(marcador):
                for idx, item in enumerate(st.session_state.dados_sessao[marcador]):
                    col1, col2 = st.columns([0.9, 0.1])
                    col1.caption(f"📄 {item['name']}")
                    if col2.button("🗑️", key=f"del_{marcador}_{idx}"):
                        st.session_state.dados_sessao[marcador].pop(idx); st.rerun()

# --- GERAÇÃO FINAL ---
if st.button("FINALIZAR E GERAR RELATÓRIO CACHOEIRA", type="primary", key="btn_finalizar"):
    try:
        with st.spinner("Processando indicadores e gerando documento..."):
            with tempfile.TemporaryDirectory() as tmp:
                doc = DocxTemplate("template-cachoeira.docx")
                
                # REGRAS DE SOMA - CONTRATO CACHOEIRA
                h_total_saida = sum([st.session_state.get(k, 0) for k in ["H_ALTA", "H_TRANSF_MAIOR", "H_TRANSF_MENOR", "H_EVASAO", "H_OBITO_MAIOR", "H_OBITO_MENOR"]])
                h_total_transf_int = st.session_state.get("H_TRANSF_MAIOR", 0) + st.session_state.get("H_TRANSF_INT", 0)
                h_t_obito = st.session_state.get("H_OBITO_MAIOR", 0) + st.session_state.get("H_OBITO_MENOR", 0)
                h_total_ob_int = st.session_state.get("H_OB_INT", 0) + st.session_state.get("H_OBITO_MAIOR", 0)
                h_t_atend_emerg = sum([st.session_state.get(k, 0) for k in ["H_GINECO", "H_CIR_GERAL", "H_MED_CLI", "H_ORTO", "H_PED"]])
                
                parecer_keys = [k for k in FORM_KEYS if "PARECER_" in k]
                total_amb_parecer = sum([st.session_state.get(k, 0) for k in parecer_keys])
                h_t_atend_amb = sum([st.session_state.get(k, 0) for k in ["AMB_FISIO", "AMB_PSICO", "AMB_FONO", "AMB_SERV_SOC"]]) + total_amb_parecer
                
                h_t_cir_elet = sum([st.session_state.get(k, 0) for k in ["H_ELE_CIR_GER", "H_ELE_CIR_ORTO", "H_ELE_CIR_BUCO", "H_ELE_CIR_URO"]])
                h_t_cir_emerg = sum([st.session_state.get(k, 0) for k in ["H_EMERG_CIR_GER", "H_EMERG_PART_CES", "H_EMERG_VASC", "H_EMERG_URO", "H_EMERG_ORT", "H_EMERG_GINECO"]])
                h_t_exa_proc = st.session_state.get("H_EX_ENDO", 0) + st.session_state.get("H_EX_COLO", 0)
                h_t_plan_fami = sum([st.session_state.get(k, 0) for k in ["H_PF_LAQ", "H_PF_DIU", "H_PF_BIO"]])
                h_t_proc_cir = h_t_cir_elet + h_t_cir_emerg + h_t_exa_proc + h_t_plan_fami
                
                upa_t_atend_emerg = st.session_state.get("UPA_MED_CLI", 0) + st.session_state.get("UPA_MED_PED", 0)
                upa_t_exa_proc = sum([st.session_state.get(k, 0) for k in ["UPA_EX_ELETRO", "UPA_EX_LAB", "UPA_EX_RADIO"]])
                
                contexto = {k: st.session_state.get(k, 0) for k in FORM_KEYS}
                contexto.update({
                    "SISTEMA_MES_REFERENCIA": f"{st.session_state.sel_mes}/{st.session_state.sel_ano}",
                    "H_TOTAL_SAIDA": h_total_saida,
                    "H_TOTAL_TRANSF_INT": h_total_transf_int,
                    "H_T_OBITO": h_t_obito,
                    "H_TOTAL_OB_INT": h_total_ob_int,
                    "H_T_ATEND_EMERG": h_t_atend_emerg,
                    "TOTAL_AMB_PARECER": total_amb_parecer,
                    "H_T_ATEND_AMB": h_t_atend_amb,
                    "H_T_PROC_CIR": h_t_proc_cir,
                    "H_T_CIR_ELET": h_t_cir_elet,
                    "H_T_CIR_EMERG": h_t_cir_emerg,
                    "H_T_EXA_PROC": h_t_exa_proc,
                    "H_T_PLAN_FAMI": h_t_plan_fami,
                    "UPA_T_ATEND_EMERG": upa_t_atend_emerg,
                    "UPA_T_EXA_PROC": upa_t_exa_proc,
                    "UPA_T_PESQ": st.session_state.get("UPA_PESQ_INT", 0) + st.session_state.get("UPA_PESQ_RECEP", 0),
                    "H_TOTAL_SAU_PESQ": st.session_state.get("H_SAU_PESQ_INT", 0) + st.session_state.get("H_SAU_OUV_RECEP", 0)
                })

                # PROCESSAMENTO DE EVIDÊNCIAS
                for m in DIMENSOES_CAMPOS.keys():
                    imgs_word = []
                    for item in st.session_state.dados_sessao.get(m, []):
                        res = processar_item_lista(doc, item['content'], m)
                        if res: imgs_word.extend(res)
                    contexto[m] = imgs_word

                doc.render(contexto)
                docx_p = os.path.join(tmp, "relatorio.docx")
                doc.save(docx_p)
                
                st.success("✅ Relatório gerado com sucesso!")
                cd1, cd2 = st.columns(2)
                with cd1:
                    with open(docx_p, "rb") as f_w:
                        st.download_button("WORD (.docx)", f_w.read(), f"RELATORIO_CACHOEIRA_{st.session_state.sel_mes}.docx")
                with cd2:
                    try:
                        converter_para_pdf(docx_p, tmp)
                        pdf_p = os.path.join(tmp, "relatorio.pdf")
                        if os.path.exists(pdf_p):
                            with open(pdf_p, "rb") as f_p:
                                st.download_button("PDF", f_p.read(), f"RELATORIO_CACHOEIRA_{st.session_state.sel_mes}.pdf")
                    except: st.warning("Conversão PDF não disponível no ambiente atual.")

    except Exception as e:
        st.error(f"Erro Crítico: {e}")

st.caption("Desenvolvido por Leonardo Barcelos Martins")
