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
import pydeck as pdk
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
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
        "email": "pamelaalmeida5@icloud.com",
        "lat": -21.7858,
        "lon": -46.5625,
        "cidades": ["POÇOS DE CALDAS", "ANDRADAS", "GUAXUPÉ", "ITAJUBA", "POUSO ALEGRE", "VARGINHA", "TRÊS PONTAS", "TRÊS CORAÇÕES", "MACHADO", "ALFENAS"]
    },
    "Fernanda Dias Ferreira": {
        "endereco": "Rua Jorge Raimundo, 409 - Juiz de Fora/MG",
        "email": "fernandaferreira_jf@yahoo.com.br",
        "lat": -21.7642,
        "lon": -43.3496,
        "cidades": ["JUIZ DE FORA"]
    },
    "Saruete Valeska Stabile de Oliveira": {
        "endereco": "Rua José Gonçalves de Souza, 105 - São José do Rio Preto/SP",
        "email": "saruetesjrp79@gmail.com",
        "lat": -20.8113,
        "lon": -49.3758,
        "cidades": ["SÃO JOSÉ DO RIO PRETO", "MIRASSOL", "CATANDUVA"]
    },
    "Carolina Rodrigues Bruno": {
        "endereco": "Rua Doutor Bernardino de Campos - São Carlos/SP",
        "email": "Crbruno27123@gmail.com",
        "lat": -22.0175,
        "lon": -47.8908,
        "cidades": ["SÃO CARLOS", "ARARAQUARA", "MATÃO"]
    },
    "Madalla Teixeira Reis": {
        "endereco": "Rua Odilon Machado, 105 - Tocantins/MG",
        "email": "madallareis66@gmail.com",
        "lat": -21.1764,
        "lon": -43.0181,
        "cidades": ["UBÁ", "DESCOBERTO", "SÃO JOÃO NEPOMUCENO", "VIÇOSA", "TOCANTINS", "RODEIRO", "PIRAÚBA", "GUARANI", "RIO POMBA", "RIO NOVO"]
    },
    "Rodrigo Luis Adao": {
        "endereco": "Avenida Professora Edul Rangel Rabello, 405 - Ribeirão Preto/SP",
        "email": "",
        "lat": -21.2075,
        "lon": -47.7981,
        "cidades": ["RIBEIRÃO PRETO"]
    }
}

PROMOTORES = list(DADOS_PROMOTORES.keys())
SITUACOES = ['Normal', 'Férias', 'Carro Quebrado', 'Feriado', 'Atestado Médico', 'Folga', 'Falta']
NOME_PLANILHA_CLIENTES = "Cópia de clientes com cnpj corretinho novinho (1).xlsx"
EMAIL_PRINCIPAL = "benedito.bandola@minassal.com.br"
EMAIL_REMETENTE = "beneditobandola@gmail.com"
VALOR_KM_TAXA = 1.17
ARQUIVO_JSON_GERAL = "historico_km_geral.json"

# ==============================================================================
# AUXILIARES DE FORMATAÇÃO E CÁLCULOS
# ==============================================================================
def normalizar_texto(txt):
    if not txt:
        return ""
    return unicodedata.normalize('NFKD', str(txt)).encode('ASCII', 'ignore').decode('utf-8').strip().upper()

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
    start = datetime(ano, 1, 1)
    start -= timedelta(days=start.weekday())
    segunda = start + timedelta(weeks=num_semana - 1)
    domingo = segunda + timedelta(days=6)
    return segunda, domingo

def distancia_haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

def estimar_km_circuito_completo(lat_casa, lon_casa, pontos_lojas):
    if not pontos_lojas:
        return 0.0
    rota = [(lat_casa, lon_casa)] + pontos_lojas + [(lat_casa, lon_casa)]
    dist_total = 0.0
    for i in range(len(rota) - 1):
        dist_total += distancia_haversine(
            rota[i][0], rota[i][1],
            rota[i+1][0], rota[i+1][1]
        )
    return dist_total * 1.28

# ==============================================================================
# GERAÇÃO DO PDF EXECUTIVO (MODELO MINASSAL)
# ==============================================================================
def gerar_pdf_resumo_financeiro(promotor_nome, semana_num, payload_dados, caminho_pdf_saida):
    doc = SimpleDocTemplate(caminho_pdf_saida, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    
    is_teste = payload_dados.get("is_teste", False)
    titulo_sufixo = " [TESTE / SIMULAÇÃO]" if is_teste else ""

    estilo_titulo = ParagraphStyle('T', fontName='Helvetica-Bold', fontSize=14, textColor=colors.HexColor('#1B5E20'), alignment=1)
    estilo_sub = ParagraphStyle('S', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#333333'), alignment=1)
    estilo_th = ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=9, textColor=colors.white, alignment=1)
    estilo_td = ParagraphStyle('TD', fontName='Helvetica', fontSize=8.5, textColor=colors.HexColor('#222222'), alignment=1)
    estilo_td_l = ParagraphStyle('TDL', parent=estilo_td, alignment=0)

    elementos = []
    elementos.append(Paragraph(f"MINASSAL CONTROLE DE REEMBOLSO{titulo_sufixo}", estilo_titulo))
    elementos.append(Spacer(1, 4))
    elementos.append(Paragraph(f"RESUMO FINANCEIRO — Semana {semana_num} de {datetime.now().year}", estilo_sub))
    elementos.append(Spacer(1, 4))
    elementos.append(Paragraph(f"<b>Promotor(a):</b> {promotor_nome}", ParagraphStyle('P', fontName='Helvetica', fontSize=9, textColor=colors.HexColor('#444444'))))
    elementos.append(Spacer(1, 8))
    elementos.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#1B5E20'), spaceAfter=10))

    tabela_dados = [[
        Paragraph("DATA", estilo_th),
        Paragraph("SITUAÇÃO", estilo_th),
        Paragraph("KM RODADO", estilo_th),
        Paragraph("REEMBOLSO", estilo_th),
    ]]

    for d in payload_dados.get("detalhes", []):
        km_d = d.get("km", 0.0)
        reemb_d = km_d * VALOR_KM_TAXA
        tabela_dados.append([
            Paragraph(d.get("data", ""), estilo_td),
            Paragraph(d.get("sit", "Normal"), estilo_td),
            Paragraph(f"{float_para_str_br(km_d)}", estilo_td),
            Paragraph(f"R$ {float_para_str_br(reemb_d)}", estilo_td),
        ])

    t = Table(tabela_dados, colWidths=[100, 150, 110, 143])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1B5E20')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#FAFAFA')),
    ]))
    elementos.append(t)
    elementos.append(Spacer(1, 15))

    elementos.append(Paragraph("<b>DESPESAS EXTRAS REEMBOLSÁVEIS:</b>", ParagraphStyle('DE', fontName='Helvetica-Bold', fontSize=9.5, textColor=colors.HexColor('#1B5E20'))))
    elementos.append(Spacer(1, 5))

    gastos = payload_dados.get("gastos_extras", [])
    if gastos:
        tabela_gastos = [[Paragraph("DESCRIÇÃO", estilo_th), Paragraph("VALOR", estilo_th)]]
        for g in gastos:
            tabela_gastos.append([
                Paragraph(g.get("desc", ""), estilo_td_l),
                Paragraph(f"R$ {float_para_str_br(g.get('valor', 0.0))}", estilo_td),
            ])
        tg = Table(tabela_gastos, colWidths=[350, 153])
        tg.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#333333')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ]))
        elementos.append(tg)
    else:
        elementos.append(Paragraph("Nenhuma despesa extra registrada.", ParagraphStyle('NDE', fontName='Helvetica-Oblique', fontSize=8.5, textColor=colors.HexColor('#666666'))))

    elementos.append(Spacer(1, 20))

    total_geral = payload_dados.get("valor_total", 0.0)
    estilo_total = ParagraphStyle('TOT', fontName='Helvetica-Bold', fontSize=12, textColor=colors.HexColor('#1B5E20'), alignment=2)
    elementos.append(Paragraph(f"VALOR TOTAL A PAGAR: R$ {float_para_str_br(total_geral)}", estilo_total))

    doc.build(elementos)

# ==============================================================================
# FUNÇÃO DE ENVIO DE E-MAIL COM ANEXO PDF
# ==============================================================================
def enviar_email_com_pdf(promotor_nome, semana_num, intervalo, payload_dados):
    destinatario_promotor = DADOS_PROMOTORES.get(promotor_nome, {}).get("email", "")
    destinatarios = [EMAIL_PRINCIPAL]
    if destinatario_promotor:
        destinatarios.append(destinatario_promotor)

    smtp_password = st.secrets.get("SMTP_PASSWORD", "")
    if not smtp_password:
        return False, "Senha SMTP não configurada nos Secrets."

    is_teste = payload_dados.get("is_teste", False)
    sufixo_assunto = " [TESTE]" if is_teste else ""

    nome_arq_pdf = f"Resumo_Financeiro_Semana_{num_semana}_{promotor_nome.replace(' ', '_')}.pdf"
    gerar_pdf_resumo_financeiro(promotor_nome, semana_num, payload_dados, nome_arq_pdf)

    assunto = f"[Minassal KM]{sufixo_assunto} Relatório de Reembolso - Semana {semana_num} - {promotor_nome}"
    
    corpo_html = f"""
    <html>
      <body style="font-family: Arial, sans-serif; color: #333;">
        <h2 style="color: #1B5E20;">Minassal - Fechamento de KM e Reembolso {sufixo_assunto}</h2>
        <p><b>Promotor(a):</b> {promotor_nome}</p>
        <p><b>Semana de Referência:</b> Semana {semana_num} ({intervalo})</p>
        <hr/>
        <p>Segue em anexo o resumo financeiro executivo em PDF contendo o detalhamento de KM, rotas e despesas extras.</p>
        <p><b>Valor Total a Pagar: R$ {float_para_str_br(payload_dados['valor_total'])}</b></p>
        <p style="font-size: 11px; color: #777; margin-top: 20px;">E-mail automático enviado pelo sistema de Controle de KM da Minassal.</p>
      </body>
    </html>
    """

    msg = MIMEMultipart()
    msg["Subject"] = assunto
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = ", ".join(destinatarios)
    msg.attach(MIMEText(corpo_html, "html"))

    try:
        with open(nome_arq_pdf, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {nome_arq_pdf}")
        msg.attach(part)
    except Exception as e:
        return False, f"Erro ao anexar PDF: {e}"

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(EMAIL_REMETENTE, smtp_password)
        server.sendmail(EMAIL_REMETENTE, destinatarios, msg.as_string())
        server.quit()
        if os.path.exists(nome_arq_pdf):
            os.remove(nome_arq_pdf)
        return True, "E-mail com PDF enviado com sucesso!"
    except Exception as e:
        return False, str(e)

# ==============================================================================
# CARREGAMENTO DA BASE DE CLIENTES E VENDAS
# ==============================================================================
@st.cache_data(ttl=1800)
def carregar_base_cruzada():
    arqs = glob.glob("*.xlsx")
    caminho_cli = NOME_PLANILHA_CLIENTES
    if caminho_cli not in arqs and arqs:
        for a in arqs:
            if "clientes" in a.lower():
                caminho_cli = a
                break
        else:
            caminho_cli = arqs[0]

    try:
        df_cli = pd.read_excel(caminho_cli, sheet_name=0)
        df_cli = df_cli.dropna(subset=["CÓDIGO", "NOME"]).copy()
        df_cli["CÓDIGO"] = df_cli["CÓDIGO"].astype(int).astype(str).str.strip()
        df_cli["NOME"] = df_cli["NOME"].astype(str).str.strip()
        df_cli["CIDADE_RAW"] = df_cli["CIDADE"].fillna("NÃO INFORMADA").astype(str).str.strip().str.upper()
        df_cli["CIDADE_NORM"] = df_cli["CIDADE_RAW"].apply(normalizar_texto)
        df_cli["BAIRRO"] = df_cli["BAIRRO"].fillna("").astype(str).str.strip()
        df_cli["ENDEREÇO"] = df_cli["ENDEREÇO"].fillna("").astype(str).str.strip()
        df_cli["UF"] = df_cli["UF"].fillna("").astype(str).str.strip()

        def sanitizar_coord(val):
            try:
                if pd.isna(val):
                    return None
                return float(str(val).replace(",", ".").strip())
            except ValueError:
                return None

        df_cli["lat"] = df_cli["LATITUDE"].apply(sanitizar_coord)
        df_cli["lon"] = df_cli["LONGITUDE"].apply(sanitizar_coord)
    except Exception as e:
        st.error(f"Erro ao carregar cadastro de clientes: {e}")
        return pd.DataFrame()

    candidatos_cubo = [a for a in arqs if "cubo" in a.lower() or "vendas" in a.lower()]
    if not candidatos_cubo:
        return df_cli

    caminho_vendas = max(candidatos_cubo, key=os.path.getmtime)
    try:
        df_vendas = pd.read_excel(caminho_vendas, sheet_name=0)
        df_vendas = df_vendas.dropna(subset=["CLIENTE CODIGO", "TOTAL VALOR"]).copy()
        df_vendas["CLIENTE CODIGO"] = df_vendas["CLIENTE CODIGO"].astype(int).astype(str).str.strip()
        df_vendas = df_vendas[df_vendas["TOTAL VALOR"] > 0]
        vendas_resumo = df_vendas.groupby("CLIENTE CODIGO")["TOTAL VALOR"].sum().reset_index()
    except Exception:
        return df_cli

    df_merged = pd.merge(df_cli, vendas_resumo, left_on="CÓDIGO", right_on="CLIENTE CODIGO", how="inner")
    df_merged = df_merged.sort_values(by="TOTAL VALOR", ascending=False)
    return df_merged

DF_CLIENTES = carregar_base_cruzada()

MAPA_GERAL_NOMES = {}
if not DF_CLIENTES.empty:
    for _, r in DF_CLIENTES.iterrows():
        bairro_str = f" - {r['BAIRRO']}" if r['BAIRRO'] else ""
        cidade_str = f" ({r['CIDADE_RAW']}/{r['UF']})" if r['CIDADE_RAW'] else ""
        MAPA_GERAL_NOMES[r["CÓDIGO"]] = f"{r['NOME']}{bairro_str}{cidade_str}"

# ==============================================================================
# INTEGRAÇÃO COM GITHUB (JSON CENTRALIZADO)
# ==============================================================================
def get_github_credentials():
    try:
        return (
            st.secrets["GITHUB_TOKEN"],
            st.secrets["GITHUB_REPO"],
            st.secrets.get("GITHUB_BRANCH", "main")
        )
    except Exception:
        return None, None, None

def carregar_base_historico_github():
    token, repo, branch = get_github_credentials()
    if not token or not repo:
        return {}, None

    url = f"https://api.github.com/repos/{repo}/contents/{ARQUIVO_JSON_GERAL}?ref={branch}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github.v3+json"}
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            data = response.json()
            conteudo_base64 = data.get("content", "")
            sha = data.get("sha", "")
            conteudo_json = json.loads(base64.b64decode(conteudo_base64).decode('utf-8'))
            return conteudo_json, sha
    except requests.RequestException:
        pass
    return {}, None

def salvar_base_historico_github(historico_dict, sha_existente=None, mensagem_commit="Atualização geral de KM"):
    token, repo, branch = get_github_credentials()
    if not token or not repo:
        st.error("Credenciais do GitHub não configuradas nos Secrets do Streamlit!")
        return False

    url = f"https://api.github.com/repos/{repo}/contents/{ARQUIVO_JSON_GERAL}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github.v3+json"}

    conteudo_json_str = json.dumps(historico_dict, indent=4, ensure_ascii=False)
    conteudo_base64 = base64.b64encode(conteudo_json_str.encode('utf-8')).decode('utf-8')

    payload = {
        "message": mensagem_commit,
        "content": conteudo_base64,
        "branch": branch
    }
    if sha_existente:
        payload["sha"] = sha_existente

    try:
        response = requests.put(url, headers=headers, json=payload, timeout=10)
        return response.status_code in [200, 201]
    except requests.RequestException as e:
        st.error(f"Erro ao salvar dados no GitHub: {e}")
        return False

HISTORICO_GERAL, SHA_GERAL = carregar_base_historico_github()

# ==============================================================================
# TELA DE IDENTIFICAÇÃO (COM OPÇÃO DE ÁREA DE TESTES)
# ==============================================================================
if "usuario_ativo" not in st.session_state:
    st.session_state.usuario_ativo = None

if not st.session_state.usuario_ativo:
    st.markdown("<h1 style='text-align: center; color: #FDD818 !important;'># ACESSO DE PROMOTORES</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #888;'>SELECIONE SEU PERFIL OU ENTRE NA ÁREA DE TESTES</p>", unsafe_allow_html=True)
    
    opcoes_acesso = ["-- Selecione seu perfil ou área --", "🧪 [ÁREA DE TESTES / SIMULAÇÃO]"] + PROMOTORES
    escolha_promotor = st.selectbox("QUEM É VOCÊ?", options=opcoes_acesso)
    
    if st.button("ACESSAR SISTEMA ➔", type="primary"):
        if escolha_promotor != "-- Selecione seu perfil ou área --":
            st.session_state.usuario_ativo = escolha_promotor
            st.rerun()
        else:
            st.warning("Por favor, selecione uma opção antes de prosseguir.")
    st.stop()

promotor_sel = st.session_state.usuario_ativo
is_area_teste = ("ÁREA DE TESTES" in promotor_sel)

if is_area_teste:
    st.markdown("### 🧪 ÁREA DE TESTES E SIMULAÇÃO")
    st.info("Escolha abaixo o promotor (ou todos em lote), informe a semana e clique para gerar e disparar os testes de e-mail com PDF.")
    
    escolha_teste_promotor = st.selectbox("ESCOLHA O PROMOTOR PARA O TESTE:", options=["🔄 Todos os Promotores (Lote)"] + PROMOTORES)
    num_semana_teste = st.number_input("Nº DA SEMANA PARA O TESTE:", min_value=1, max_value=53, value=int(datetime.now().isocalendar()[1]))
    
    if st.button("⚡ GERAR E DISPARAR TESTE(S) AGORA"):
        promotores_alvo = PROMOTORES if escolha_teste_promotor == "🔄 Todos os Promotores (Lote)" else [escolha_teste_promotor]
        seg_t, dom_t = calcular_intervalo_semana(num_semana_teste)
        int_str_t = f"{seg_t.strftime('%d/%m')} a {dom_t.strftime('%d/%m')}"

        sucessos = 0
        for p_nome in promotores_alvo:
            p_dados = DADOS_PROMOTORES[p_nome]
            cidades_p = [normalizar_texto(c) for c in p_dados.get("cidades", [])]

            dias_semana = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
            detalhes_ficticios = []
            km_tot_fict = 0.0
            km_base = 100.0

            for idx_d, d_nome in enumerate(dias_semana):
                dt_d = (seg_t + timedelta(days=idx_d)).strftime("%d/%m")
                sit_f = "Normal" if idx_d not in [5, 6] else "Folga"
                cli_f = []
                km_d_f = 0.0
                kmi_f = km_base
                kmf_f = km_base

                if sit_f == "Normal" and not DF_CLIENTES.empty:
                    df_prom = DF_CLIENTES[DF_CLIENTES["CIDADE_NORM"].isin(cidades_p)]
                    if df_prom.empty:
                        df_prom = DF_CLIENTES
                    amostra = df_prom.sample(n=min(3, len(df_prom)))
                    cli_f = amostra["CÓDIGO"].tolist()
                    km_d_f = round(random.uniform(45.0, 95.0), 1)
                    kmf_f = kmi_f + km_d_f
                    km_base = kmf_f

                detalhes_ficticios.append({
                    "dia": d_nome,
                    "data": dt_d,
                    "km": km_d_f,
                    "sit": sit_f,
                    "clientes": cli_f,
                    "leitura": True,
                    "km_ini": kmi_f,
                    "km_fim": kmf_f
                })
                km_tot_fict += km_d_f

            gastos_fict = [{"desc": "Almoço de Teste", "valor": 42.50}, {"desc": "Estacionamento", "valor": 15.00}]
            v_km_fict = km_tot_fict * VALOR_KM_TAXA
            v_ext_fict = sum(g["valor"] for g in gastos_fict)
            v_tot_fict = v_km_fict + v_ext_fict

            payload_teste = {
                "id": int(datetime.now().timestamp()),
                "semana_ref": str(num_semana_teste),
                "intervalo_datas": int_str_t,
                "promotor": p_nome,
                "status": "FINALIZADO",
                "is_teste": True,
                "km_total": km_tot_fict,
                "valor_km": v_km_fict,
                "valor_extras": v_ext_fict,
                "valor_total": v_tot_fict,
                "gastos_extras": gastos_fict,
                "detalhes": detalhes_ficticios
            }

            chave_reg_t = f"{p_nome}_S{num_semana_teste}"
            HISTORICO_GERAL[chave_reg_t] = payload_teste
            ok_mail, _ = enviar_email_com_pdf(p_nome, num_semana_teste, int_str_t, payload_teste)
            if ok_mail:
                sucessos += 1

        ok_salvar = salvar_base_historico_github(HISTORICO_GERAL, sha_existente=SHA_GERAL, mensagem_commit=f"Testes em lote S{num_semana_teste}")
        if ok_salvar and sucessos > 0:
            st.success(f"✅ Testes executados com sucesso! {sucessos} relatório(s) gerado(s) e e-mail(s) disparado(s).")
            st.balloons()
        else:
            st.warning("Houve falha ao salvar no GitHub ou disparar os e-mails.")

    st.write("")
    if st.button("⬅️ VOLTAR À SELEÇÃO DE PERFIL"):
        st.session_state.usuario_ativo = None
        st.rerun()
    st.stop()

# ==============================================================================
# FLUXO NORMAL DO PROMOTOR
# ==============================================================================
dados_promotor_atual = DADOS_PROMOTORES.get(promotor_sel, {})
cidades_definidas = dados_promotor_atual.get("cidades", [])
cidades_norm_promotor = [normalizar_texto(c) for c in cidades_definidas]

if not DF_CLIENTES.empty:
    todas_cidades_norm = sorted(list(DF_CLIENTES["CIDADE_NORM"].unique()))
    if cidades_norm_promotor:
        cidades_disponiveis_promotor = [c for c in cidades_norm_promotor if c in todas_cidades_norm]
    else:
        cidades_disponiveis_promotor = todas_cidades_norm
else:
    cidades_disponiveis_promotor = []

col_tit, col_sair = st.columns([3, 1])
with col_tit:
    st.markdown("<h2 style='color:#FDD818 !important; margin:0;'># CONTROLE DE KM</h2>", unsafe_allow_html=True)
with col_sair:
    st.write("")
    if st.button("TROCAR 🔄"):
        st.session_state.usuario_ativo = None
        st.session_state.clear()
        st.rerun()

st.markdown(f"**PROMOTOR(A):** {promotor_sel}")
st.caption(f"🏠 {dados_promotor_atual.get('endereco', 'Não cadastrado')}")

semana_atual_default = int(datetime.now().isocalendar()[1])
num_semana = st.number_input("Nº DA SEMANA:", min_value=1, max_value=53, value=semana_atual_default)

segunda, domingo = calcular_intervalo_semana(num_semana)
intervalo_str = f"{segunda.strftime('%d/%m')} a {domingo.strftime('%d/%m')}"
st.markdown(f"<div style='padding:6px 12px; background:#141414; border-left:4px solid #FDD818; margin-bottom:15px;'>📅 <b>PERÍODO:</b> Semana {num_semana} ({intervalo_str})</div>", unsafe_allow_html=True)

chave_registro = f"{promotor_sel}_S{num_semana}"
dados_salvos = HISTORICO_GERAL.get(chave_registro, None)

with st.expander("🛠️ GERENCIAR OU APAGAR LANÇAMENTOS", expanded=False):
    st.markdown("Selecione um lançamento cadastrado no histórico geral para excluí-lo:")
    semanas_cadastradas = list(HISTORICO_GERAL.keys())
    if semanas_cadastradas:
        chave_para_apagar = st.selectbox("Lançamento cadastrado:", semanas_cadastradas)
        if st.button("🗑️ APAGAR ESTE LANÇAMENTO SELECIONADO"):
            del HISTORICO_GERAL[chave_para_apagar]
            sucesso_del = salvar_base_historico_github(
                HISTORICO_GERAL, 
                sha_existente=SHA_GERAL, 
                mensagem_commit=f"Removido lançamento {chave_para_apagar}"
            )
            if sucesso_del:
                st.success(f"Lançamento {chave_para_apagar} apagado com sucesso!")
                st.rerun()
    else:
        st.info("Nenhum lançamento no histórico.")

st.divider()

esta_finalizado = False
if dados_salvos and dados_salvos.get("status") == "FINALIZADO":
    esta_finalizado = True
    is_t = dados_salvos.get("is_teste", False)
    txt_t = " [TESTE]" if is_t else ""
    st.success(f"🔒 **SEMANA FINALIZADA E TRANSMITIDA{txt_t}.** Todos os campos estão bloqueados.")
    if st.button("🔓 REABRIR PARA CORREÇÃO"):
        dados_salvos["status"] = "RASCUNHO"
        HISTORICO_GERAL[chave_registro] = dados_salvos
        sucesso_reabrir = salvar_base_historico_github(
            HISTORICO_GERAL,
            sha_existente=SHA_GERAL,
            mensagem_commit=f"Reaberto para edição S{num_semana} - {promotor_sel}"
        )
        if sucesso_reabrir:
            st.success("Semana destravada com sucesso!")
            st.rerun()
elif dados_salvos:
    st.warning("📝 Rascunho salvo em aberto. Edição liberada.")

# ==============================================================================
# REGISTROS DIÁRIOS
# ==============================================================================
dias_semana = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
detalhes_dias = []
km_total_calculado = 0.0
km_fim_anterior = 0.0

mapa_dados_salvos = {}
if dados_salvos and "detalhes" in dados_salvos:
    mapa_dados_salvos = {d["dia"]: d for d in dados_salvos["detalhes"]}

st.markdown("### 📋 REGISTROS DIÁRIOS")

for i, dia_nome in enumerate(dias_semana):
    data_dia = (segunda + timedelta(days=i)).strftime("%d/%m")
    dados_dia_salvo = mapa_dados_salvos.get(dia_nome, {})

    with st.expander(f"📌 {dia_nome.upper()} ({data_dia})", expanded=(i == 0 or bool(dados_dia_salvo))):
        col_sit, col_lei = st.columns([2, 1])
        with col_sit:
            sit_default = dados_dia_salvo.get("sit", "Normal")
            idx_sit = SITUACOES.index(sit_default) if sit_default in SITUACOES else 0
            sit_sel = st.selectbox(
                f"Situação ({dia_nome})", 
                SITUACOES, 
                index=idx_sit, 
                disabled=esta_finalizado,
                key=f"sit_{dia_nome}"
            )

        with col_lei:
            lei_default = dados_dia_salvo.get("leitura", False)
            lei_sel = st.checkbox(
                "Leituras?", 
                value=lei_default, 
                disabled=esta_finalizado,
                key=f"lei_{dia_nome}"
            )

        dia_eh_normal = (sit_sel == "Normal")
        clientes_dia_selecionados = []
        df_atendidos = pd.DataFrame()
        km_sugerido_circuito = 0.0

        if dia_eh_normal:
            st.markdown("#### 🏬 LOJAS ATENDIDAS")
            cods_salvos_dia = [str(c) for c in dados_dia_salvo.get("clientes", [])]
            session_key_lojas = f"lojas_selecionadas_{dia_nome}"
            if session_key_lojas not in st.session_state:
                st.session_state[session_key_lojas] = cods_salvos_dia

            if not DF_CLIENTES.empty:
                for cid_norm in cidades_disponiveis_promotor:
                    df_cidade = DF_CLIENTES[DF_CLIENTES["CIDADE_NORM"] == cid_norm]
                    if df_cidade.empty:
                        continue
                    
                    nome_cidade_exibicao = df_cidade["CIDADE_RAW"].iloc[0]
                    codigos_da_cidade = df_cidade["CÓDIGO"].tolist()
                    defaults_cidade = [c for c in st.session_state[session_key_lojas] if c in codigos_da_cidade]

                    with st.expander(f"🏙️ {nome_cidade_exibicao} ({len(df_cidade)} lojas)", expanded=bool(defaults_cidade)):
                        escolhidos_cid = st.multiselect(
                            f"Lojas em {nome_cidade_exibicao}:",
                            options=codigos_da_cidade,
                            default=defaults_cidade,
                            disabled=esta_finalizado,
                            format_func=lambda cod: MAPA_GERAL_NOMES.get(cod, f"Cód: {cod}"),
                            key=f"cli_{dia_nome}_{cid_norm}"
                        )
                        clientes_dia_selecionados.extend(escolhidos_cid)

                with st.expander("🌐 Outra Cidade / Fora da Rota", expanded=False):
                    cidades_fora = [c for c in todas_cidades_norm if c not in cidades_disponiveis_promotor]
                    cid_extra = st.selectbox(
                        f"Escolha a cidade ({dia_nome}):", 
                        ["-- Selecione --"] + cidades_fora, 
                        disabled=esta_finalizado,
                        key=f"extra_cid_{dia_nome}"
                    )
                    if cid_extra != "-- Selecione --":
                        df_extra = DF_CLIENTES[DF_CLIENTES["CIDADE_NORM"] == cid_extra]
                        cods_extra = df_extra["CÓDIGO"].tolist()
                        defaults_extra = [c for c in st.session_state[session_key_lojas] if c in cods_extra]
                        escolhidos_extra = st.multiselect(
                            f"Lojas em {cid_extra}:",
                            options=cods_extra,
                            default=defaults_extra,
                            disabled=esta_finalizado,
                            format_func=lambda cod: MAPA_GERAL_NOMES.get(cod, f"Cód: {cod}"),
                            key=f"cli_extra_{dia_nome}"
                        )
                        clientes_dia_selecionados.extend(escolhidos_extra)

            clientes_dia_selecionados = list(dict.fromkeys(clientes_dia_selecionados))
            st.session_state[session_key_lojas] = clientes_dia_selecionados

            if clientes_dia_selecionados and not DF_CLIENTES.empty and "lat" in dados_promotor_atual:
                df_atendidos = DF_CLIENTES[DF_CLIENTES["CÓDIGO"].isin(clientes_dia_selecionados)]
                df_com_coord = df_atendidos.dropna(subset=["lat", "lon"])
                if not df_com_coord.empty:
                    pontos_visitas = list(zip(df_com_coord["lat"], df_com_coord["lon"]))
                    km_sugerido_circuito = round(
                        estimar_km_circuito_completo(dados_promotor_atual["lat"], dados_promotor_atual["lon"], pontos_visitas),
                        1
                    )
        else:
            st.markdown(f"<div style='color:#FDD818; padding:8px 0;'>⚠️ Dia registrado como <b>{sit_sel}</b>. Lojas desativadas.</div>", unsafe_allow_html=True)
            clientes_dia_selecionados = []

        col_kmi, col_kmf = st.columns(2)
        with col_kmi:
            def_kmi = str_br_para_float(dados_dia_salvo.get("km_ini", km_fim_anterior))
            km_ini_str = st.text_input(
                f"KM Inicial ({dia_nome})", 
                value=float_para_str_br(def_kmi).replace(",00", ""), 
                disabled=esta_finalizado or (not dia_eh_normal),
                key=f"kmi_{dia_nome}"
            )
            km_ini = str_br_para_float(km_ini_str)

        with col_kmf:
            if dia_eh_normal:
                val_salvo_kmf = dados_dia_salvo.get("km_fim", None)
                if val_salvo_kmf is not None and str(val_salvo_kmf) not in ["0", "0.0", ""]:
                    def_kmf = str_br_para_float(val_salvo_kmf)
                elif km_sugerido_circuito > 0.0 and km_ini > 0.0:
                    def_kmf = km_ini + km_sugerido_circuito
                else:
                    def_kmf = def_kmi
            else:
                def_kmf = km_ini

            km_fim_str = st.text_input(
                f"KM Final ({dia_nome})", 
                value=float_para_str_br(def_kmf).replace(",00", ""), 
                disabled=esta_finalizado or (not dia_eh_normal),
                key=f"kmf_{dia_nome}"
            )
            km_fim = str_br_para_float(km_fim_str)

        if dia_eh_normal and km_sugerido_circuito > 0.0 and not esta_finalizado:
            st.markdown(f"<div style='color:#888; font-size:13px;'>💡 <b>Sugestão de rota:</b> ~{float_para_str_br(km_sugerido_circuito)} km (Ida e Volta). Editável livremente.</div>", unsafe_allow_html=True)

        km_dia = 0.0
        if dia_eh_normal and km_fim > 0.0:
            if km_fim < km_ini:
                st.error("⚠️ KM Final não pode ser menor que o KM Inicial!")
            else:
                km_dia = km_fim - km_ini
                km_fim_anterior = km_fim
                st.markdown(f"<span style='color:#FDD818; font-weight:bold;'>🚘 KM Rodado: {float_para_str_br(km_dia)} km</span>", unsafe_allow_html=True)
        elif not dia_eh_normal:
            km_fim_anterior = km_ini

        km_total_calculado += km_dia

        if dia_eh_normal and clientes_dia_selecionados and not df_atendidos.empty:
            with st.expander(f"📍 Lojas Selecionadas ({len(clientes_dia_selecionados)})", expanded=True):
                for _, row in df_atendidos.iterrows():
                    st.markdown(f"• **{row['NOME']}** \n 🏠 {row['ENDEREÇO']} - {row['BAIRRO']}, {row['CIDADE_RAW']}/{row['UF']}")

            df_coords = df_atendidos.dropna(subset=["lat", "lon"]).copy()

            if not df_coords.empty and "lat" in dados_promotor_atual:
                lat_casa = dados_promotor_atual["lat"]
                lon_casa = dados_promotor_atual["lon"]

                camada_casa = pdk.Layer(
                    "ScatterplotLayer",
                    data=pd.DataFrame([{"lat": lat_casa, "lon": lon_casa}]),
                    get_position=["lon", "lat"],
                    get_color=[0, 140, 255, 230],
                    get_radius=800,
                    pickable=True
                )

                camada_lojas = pdk.Layer(
                    "ScatterplotLayer",
                    data=df_coords,
                    get_position=["lon", "lat"],
                    get_color=[253, 216, 24, 230], 
                    get_radius=500,
                    pickable=True
                )

                linhas_circuito = []
                coords_lista = list(zip(df_coords["lat"], df_coords["lon"]))
                rota_pts = [(lat_casa, lon_casa)] + coords_lista + [(lat_casa, lon_casa)]
                for idx_pt in range(len(rota_pts) - 1):
                    linhas_circuito.append({
                        "origem": [rota_pts[idx_pt][1], rota_pts[idx_pt][0]],
                        "destino": [rota_pts[idx_pt+1][1], rota_pts[idx_pt+1][0]]
                    })

                camada_linhas = pdk.Layer(
                    "LineLayer",
                    data=pd.DataFrame(linhas_circuito),
                    get_source_position="origem",
                    get_target_position="destino",
                    get_color=[255, 255, 255, 160], 
                    get_width=3
                )

                viewport = pdk.ViewState(latitude=lat_casa, longitude=lon_casa, zoom=10, pitch=0)
                st.pydeck_chart(pdk.Deck(layers=[camada_linhas, camada_casa, camada_lojas], initial_view_state=viewport, map_style="dark"))

        detalhes_dias.append({
            "dia": dia_nome,
            "data": data_dia,
            "km": km_dia,
            "sit": sit_sel,
            "clientes": clientes_dia_selecionados,
            "leitura": lei_sel,
            "km_ini": km_ini,
            "km_fim": km_fim
        })

# ==============================================================================
# GASTOS EXTRAS COM BOTÃO DINÂMICO "ADICIONAR DESPESA"
# ==============================================================================
st.divider()
st.markdown("### 💰 GASTOS EXTRAS")

# Inicializar contador de gastos extras no session_state se não existir
if "num_gastos_extras" not in st.session_state:
    gastos_salvos_default = dados_salvos.get("gastos_extras", []) if dados_salvos else []
    st.session_state.num_gastos_extras = max(1, len(gastos_salvos_default))

gastos_extras = []
for idx in range(st.session_state.num_gastos_extras):
    gastos_salvos_default = dados_salvos.get("gastos_extras", []) if dados_salvos else []
    g_item = gastos_salvos_default[idx] if idx < len(gastos_salvos_default) else {}
    
    col_desc, col_val, col_del = st.columns([3, 2, 0.8])
    with col_desc:
        g_desc = st.text_input(
            f"Despesa #{idx+1}", 
            value=g_item.get("desc", ""), 
            placeholder="Ex: Estacionamento, Refeição...", 
            disabled=esta_finalizado,
            key=f"gdesc_{idx}"
        )
    with col_val:
        v_salvo_num = g_item.get("valor", 0.0)
        v_salvo_txt = float_para_str_br(v_salvo_num) if v_salvo_num > 0 else ""
        g_val_txt = st.text_input(
            f"Valor R$ #{idx+1}", 
            value=v_salvo_txt, 
            placeholder="0,00", 
            disabled=esta_finalizado,
            key=f"gval_{idx}"
        )
        v_float = str_br_para_float(g_val_txt)

    if g_desc.strip() and v_float > 0:
        gastos_extras.append({"desc": g_desc.strip(), "valor": v_float})

if not esta_finalizado:
    if st.button("➕ ADICIONAR OUTRA DESPESA EXTRA"):
        st.session_state.num_gastos_extras += 1
        st.rerun()

# ==============================================================================
# RESUMO FINANCEIRO
# ==============================================================================
st.divider()
st.markdown("### 📊 FECHAMENTO DA SEMANA")

valor_total_km = km_total_calculado * VALOR_KM_TAXA
valor_extras_total = sum(g["valor"] for g in gastos_extras)
valor_total_reembolso = valor_total_km + valor_extras_total

c_km, c_ext, c_tot = st.columns(3)
c_km.metric("REEMBOLSO KM", f"R$ {float_para_str_br(valor_total_km)}")
c_ext.metric("GASTOS EXTRAS", f"R$ {float_para_str_br(valor_extras_total)}")
c_tot.metric("TOTAL A RECEBER", f"R$ {float_para_str_br(valor_total_reembolso)}")

def construir_payload(status_envio):
    return {
        "id": int(datetime.now().timestamp()),
        "semana_ref": str(num_semana),
        "intervalo_datas": intervalo_str,
        "promotor": promotor_sel,
        "status": status_envio,
        "is_teste": False,
        "km_total": km_total_calculado,
        "valor_km": valor_total_km,
        "valor_extras": valor_extras_total,
        "valor_total": valor_total_reembolso,
        "gastos_extras": gastos_extras,
        "detalhes": detalhes_dias
    }

st.write("")
if not esta_finalizado:
    col_btn1, col_btn2 = st.columns(2)

    with col_btn1:
        if st.button("SALVAR RASCUNHO 💾"):
            payload = construir_payload("RASCUNHO")
            HISTORICO_GERAL[chave_registro] = payload
            sucesso = salvar_base_historico_github(
                HISTORICO_GERAL, 
                sha_existente=SHA_GERAL,
                mensagem_commit=f"Rascunho S{num_semana} - {promotor_sel}"
            )
            if sucesso:
                st.success("Rascunho salvo com sucesso!")
                st.rerun()

    with col_btn2:
        if st.button("FINALIZAR SEMANA 🚀", type="primary"):
            payload = construir_payload("FINALIZADO")
            HISTORICO_GERAL[chave_registro] = payload
            sucesso = salvar_base_historico_github(
                HISTORICO_GERAL, 
                sha_existental=SHA_GERAL,
                mensagem_commit=f"FINALIZADO S{num_semana} - {promotor_sel}"
            )
            if sucesso:
                ok_email, msg_email = enviar_email_com_pdf(promotor_sel, num_semana, intervalo_str, payload)
                if ok_email:
                    st.success("Semana finalizada, bloqueada e e-mail com PDF executivo enviado com sucesso!")
                else:
                    st.warning(f"Semana finalizada no GitHub, mas houve um erro ao enviar o e-mail: {msg_email}")
                st.balloons()
                st.rerun()
else:
    st.info("ℹ️ Para realizar qualquer edição, clique em '🔓 REABRIR PARA CORREÇÃO' no topo da tela.")
