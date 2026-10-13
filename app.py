import streamlit as st
import pandas as pd
import requests
import json
import base64
import os
import glob
import math
import random
import unicodedata
import smtplib
from io import BytesIO
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA (TEMA DARK + AMARELO OURO)
# ==============================================================================
st.set_page_config(
    page_title="Minassal - Controle de KM",
    page_icon="🚗",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    .stApp {
        background-color: #050505;
        color: #EDEDED;
    }
    h1, h2, h3, h4 {
        color: #FFFFFF !important;
        font-weight: 900 !important;
        letter-spacing: 0.5px;
    }
    .stButton>button {
        width: 100%;
        background-color: #FDD818 !important;
        color: #000000 !important;
        font-weight: 900 !important;
        font-size: 15px !important;
        text-transform: uppercase;
        letter-spacing: 1px;
        border: none !important;
        border-radius: 4px !important;
        padding: 10px 18px !important;
        transition: all 0.2s ease-in-out;
    }
    .stButton>button:hover {
        background-color: #FFE24A !important;
        color: #000000 !important;
        transform: translateY(-1px);
    }
    div[data-testid="stExpander"] {
        background-color: #121212 !important;
        border: 1px solid #242424 !important;
        border-radius: 6px !important;
        margin-bottom: 12px;
    }
    div[data-testid="stExpander"] summary {
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }
    div[data-testid="stExpander"] summary:hover {
        color: #FDD818 !important;
    }
    .stTextInput input, .stSelectbox [data-baseweb="select"] {
        background-color: #1A1A1A !important;
        color: #FFFFFF !important;
        border: 1px solid #333333 !important;
        border-radius: 4px !important;
    }
    span[data-baseweb="tag"] {
        background-color: #FDD818 !important;
        color: #000000 !important;
        font-weight: bold !important;
    }
    div[data-testid="stMetricValue"] {
        color: #FDD818 !important;
        font-weight: 900 !important;
    }
    div[data-testid="stMetricLabel"] {
        color: #A0A0A0 !important;
        text-transform: uppercase;
        font-size: 12px !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# DADOS DOS PROMOTORES E E-MAILS
# ==============================================================================
DADOS_PROMOTORES = {
    "Pamela Camila de Almeida Alexandrino": {
        "endereco": "Rua Doutor Rowilson Flora, 753 - Poços de Caldas/MG",
        "email": "pamelaalmeida5@icloud.com"
    },
    "Fernanda Dias Ferreira": {
        "endereco": "Rua Jorge Raimundo, 409 - Juiz de Fora/MG",
        "email": "fernandaferreira_jf@yahoo.com.br"
    },
    "Saruete Valeska Stabile de Oliveira": {
        "endereco": "Rua José Gonçalves de Souza, 105 - São José do Rio Preto/SP",
        "email": "saruetesjrp79@gmail.com"
    },
    "Carolina Rodrigues Bruno": {
        "endereco": "Rua Doutor Bernardino de Campos - São Carlos/SP",
        "email": "Crbruno27123@gmail.com"
    },
    "Madalla Teixeira Reis": {
        "endereco": "Rua Odilon Machado, 105 - Tocantins/MG",
        "email": "madallareis66@gmail.com"
    }
}

PROMOTORES = list(DADOS_PROMOTORES.keys())
SITUACOES = ['Normal', 'Férias', 'Carro Quebrado', 'Feriado', 'Atestado Médico', 'Folga', 'Falta']
EMAIL_PRINCIPAL = "benedito.bandola@minassal.com.br"
EMAIL_ADRIANA = "adriana@minassal.com.br"
EMAIL_REMETENTE = "beneditobandola@gmail.com"
VALOR_KM_TAXA = 1.17
ARQUIVO_JSON_GERAL = "historico_km_geral.json"
NOME_LOGOTIPO = "MINASSAL_LOGOS-04.png"

# ==============================================================================
# AUXILIARES DE FORMATAÇÃO E CÁLCULOS (CALENDÁRIO EXATO ISO)
# ==============================================================================
def str_br_para_float(val):
    if not val:
        return 0.0
    s = str(val).strip().replace("R$", "").strip()
    if not s:
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0

def float_para_str_br(val):
    try:
        return f"{float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "0,00"

def calcular_intervalo_semana(num_semana, ano=None):
    if ano is None:
        ano = datetime.now().year
    primeiro_dia_ano = datetime(ano, 1, 4)
    inicio_ano_iso = primeiro_dia_ano - timedelta(days=primeiro_dia_ano.weekday())
    segunda = inicio_ano_iso + timedelta(weeks=num_semana - 1)
    domingo = segunda + timedelta(days=6)
    return segunda, domingo

def obter_semana_padrao_inteligente():
    agora = datetime.now()
    semana_padrao = agora.isocalendar()[1]
    if agora.weekday() == 0 and agora.hour < 12:
        semana_padrao -= 1
        if semana_padrao < 1:
            semana_padrao = 52
    return semana_padrao

# ==============================================================================
# RODAPÉ DE PROTEÇÃO E AUTORIA
# ==============================================================================
def adicionar_rodape(canvas, doc):
    canvas.saveState()
    canvas.setFont('Helvetica', 8)
    canvas.setFillColor(colors.HexColor('#666666'))
    texto_rodape = "Desenvolvido por Benedito Bandola"
    canvas.drawRightString(A4[0] - 30, 15, texto_rodape)
    canvas.restoreState()

# ==============================================================================
# GERAÇÃO DO PDF EXECUTIVO (COM LINHAS DE TOTAIS E OBSERVAÇÕES)
# ==============================================================================
def gerar_pdf_resumo_financeiro(promotor_nome, semana_num, payload_dados, caminho_pdf_saida):
    doc = SimpleDocTemplate(
        caminho_pdf_saida, 
        pagesize=A4, 
        rightMargin=30, 
        leftMargin=30, 
        topMargin=30, 
        bottomMargin=30
    )
    
    is_teste = payload_dados.get("is_teste", False)
    titulo_sufixo = " [TESTE / SIMULAÇÃO]" if is_teste else ""

    estilo_titulo = ParagraphStyle('T', fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor('#0B2545'), alignment=0)
    estilo_sub = ParagraphStyle('S', fontName='Helvetica-Bold', fontSize=9.5, textColor=colors.HexColor('#333333'), alignment=0)
    estilo_th = ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=9, textColor=colors.white, alignment=1)
    estilo_td = ParagraphStyle('TD', fontName='Helvetica', fontSize=8.5, textColor=colors.HexColor('#222222'), alignment=1)
    estilo_td_l = ParagraphStyle('TDL', parent=estilo_td, alignment=0)
    estilo_th_tot = ParagraphStyle('THTOT', fontName='Helvetica-Bold', fontSize=8.5, textColor=colors.HexColor('#134074'), alignment=1)

    elementos = []

    col_dir_elementos = [
        Paragraph(f"<b>MINASSAL CONTROLE DE REEMBOLSO{titulo_sufixo}</b>", estilo_titulo),
        Spacer(1, 3),
        Paragraph(f"RESUMO FINANCEIRO — Semana {semana_num} de {datetime.now().year}", estilo_sub),
        Spacer(1, 3),
        Paragraph(f"<b>Promotor(a):</b> {promotor_nome}", ParagraphStyle('P', fontName='Helvetica', fontSize=8.5, textColor=colors.HexColor('#444444')))
    ]

    if os.path.exists(NOME_LOGOTIPO):
        try:
            logo = Image(NOME_LOGOTIPO, width=110, height=45, preserveAspectRatio=True)
            logo.hAlign = 'LEFT'
            t_cabecalho = Table([[logo, col_dir_elementos]], colWidths=[120, 414])
        except Exception:
            t_cabecalho = Table([["", col_dir_elementos]], colWidths=[120, 414])
    else:
        t_cabecalho = Table([["", col_dir_elementos]], colWidths=[120, 414])

    t_cabecalho.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (0,0), 'LEFT'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))

    elementos.append(t_cabecalho)
    elementos.append(Spacer(1, 8))
    elementos.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#134074'), spaceAfter=10))

    tabela_dados = [[
        Paragraph("DATA", estilo_th),
        Paragraph("SITUAÇÃO", estilo_th),
        Paragraph("KM RODADO", estilo_th),
        Paragraph("REEMBOLSO", estilo_th),
    ]]

    soma_km_total = 0.0
    soma_reemb_total = 0.0

    for d in payload_dados.get("detalhes", []):
        km_d = d.get("km", 0.0)
        reemb_d = km_d * VALOR_KM_TAXA
        soma_km_total += km_d
        soma_reemb_total += reemb_d
        tabela_dados.append([
            Paragraph(d.get("data", ""), estilo_td),
            Paragraph(d.get("sit", "Normal"), estilo_td),
            Paragraph(f"{float_para_str_br(km_d)}", estilo_td),
            Paragraph(f"R$ {float_para_str_br(reemb_d)}", estilo_td),
        ])

    tabela_dados.append([
        Paragraph("<b>TOTAL GERAL</b>", estilo_th_tot),
        Paragraph("-", estilo_th_tot),
        Paragraph(f"<b>{float_para_str_br(soma_km_total)} km</b>", estilo_th_tot),
        Paragraph(f"<b>R$ {float_para_str_br(soma_reemb_total)}</b>", estilo_th_tot),
    ])

    t = Table(tabela_dados, colWidths=[100, 150, 110, 143])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#134074')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ('BACKGROUND', (0, 1), (-1, -2), colors.HexColor('#FAFAFA')),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#EAEAEA')),
    ]))
    elementos.append(t)
    elementos.append(Spacer(1, 15))

    elementos.append(Paragraph("<b>DESPESAS EXTRAS REEMBOLSÁVEIS:</b>", ParagraphStyle('DE', fontName='Helvetica-Bold', fontSize=9.5, textColor=colors.HexColor('#134074'))))
    elementos.append(Spacer(1, 5))

    gastos = payload_dados.get("gastos_extras", [])
    if gastos:
        tabela_gastos = [[Paragraph("DATA", estilo_th), Paragraph("DESCRIÇÃO", estilo_th), Paragraph("VALOR", estilo_th)]]
        soma_gastos_total = 0.0
        for g in gastos:
            val_g = g.get('valor', 0.0)
            soma_gastos_total += val_g
            tabela_gastos.append([
                Paragraph(g.get("data", ""), estilo_td),
                Paragraph(g.get("desc", ""), estilo_td_l),
                Paragraph(f"R$ {float_para_str_br(val_g)}", estilo_td),
            ])
        
        tabela_gastos.append([
            Paragraph("<b>TOTAL DESPESAS</b>", estilo_th_tot),
            Paragraph("-", estilo_th_tot),
            Paragraph(f"<b>R$ {float_para_str_br(soma_gastos_total)}</b>", estilo_th_tot),
        ])

        tg = Table(tabela_gastos, colWidths=[100, 250, 153])
        tg.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#134074')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
            ('BACKGROUND', (0, 1), (-1, -2), colors.HexColor('#FAFAFA')),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#EAEAEA')),
        ]))
        elementos.append(tg)
    else:
        elementos.append(Paragraph("Nenhuma despesa extra registrada.", ParagraphStyle('NDE', fontName='Helvetica-Oblique', fontSize=8.5, textColor=colors.HexColor('#666666'))))

    obs_adriana = payload_dados.get("observacao_adriana", "")
    if obs_adriana.strip():
        elementos.append(Spacer(1, 15))
        elementos.append(Paragraph("<b>OBSERVAÇÃO / JUSTIFICATIVA (BENEDITO PARA ADRIANA):</b>", ParagraphStyle('OBL', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#134074'))))
        elementos.append(Spacer(1, 3))
        elementos.append(Paragraph(obs_adriana, ParagraphStyle('OBT', fontName='Helvetica-Oblique', fontSize=8.5, textColor=colors.HexColor('#333333'))))

    elementos.append(Spacer(1, 20))

    total_geral = payload_dados.get("valor_total", 0.0)
    estilo_total = ParagraphStyle('TOT', fontName='Helvetica-Bold', fontSize=12, textColor=colors.HexColor('#134074'), alignment=2)
    elementos.append(Paragraph(f"VALOR TOTAL A PAGAR: R$ {float_para_str_br(total_geral)}", estilo_total))

    doc.build(elementos, onFirstPage=adicionar_rodape, onLaterPages=adicionar_rodape)

# ==============================================================================
# FUNÇÃO DE ENVIO DE E-MAIL COM ANEXO PDF
# ==============================================================================
def enviar_email_com_pdf(promotor_nome, semana_num, intervalo, payload_dados):
    destinatario_promotor = DADOS_PROMOTORES.get(promotor_nome, {}).get("email", "")
    destinatarios = [EMAIL_PRINCIPAL, EMAIL_ADRIANA]
    if destinatario_promotor:
        destinatarios.append(destinatario_promotor)

    smtp_password = st.secrets.get("SMTP_PASSWORD", "")
    if not smtp_password:
        return False, "Senha SMTP não configurada nos Secrets."

    is_teste = payload_dados.get("is_teste", False)
    sufixo_assunto = " [TESTE]" if is_teste else ""

    nome_arq_pdf = f"Resumo_Financeiro_Semana_{semana_num}_{promotor_nome.replace(' ', '_')}.pdf"
    gerar_pdf_resumo_financeiro(promotor_nome, semana_num, payload_dados, nome_arq_pdf)

    assunto = f"[Minassal KM]{sufixo_assunto} Relatório de Reembolso - Semana {semana_num} - {promotor_nome}"
    
    obs_adriana = payload_dados.get("observacao_adriana", "")
    bloco_obs_email = ""
    if obs_adriana.strip():
        bloco_obs_email = f"""
        <div style="background-color: #FFF9E6; border-left: 4px solid #FDD818; padding: 10px; margin: 15px 0;">
            <p style="margin: 0; font-size: 13px; color: #333;">
