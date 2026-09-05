import streamlit as st
import pandas as pd
import requests
import json
import base64
import os
import glob
import math
import unicodedata
import pydeck as pdk
from datetime import datetime, timedelta

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA (TEMA DARK + AMARELO OURO)
# ==============================================================================
st.set_page_config(
    page_title="Minassal - Controle de KM",
    page_icon="🚗",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Estilização visual inspirada na identidade de alto contraste
st.markdown("""
    <style>
    /* Fundo geral preto profundo */
    .stApp {
        background-color: #050505;
        color: #EDEDED;
    }
    
    /* Cabeçalhos em branco absoluto e caixa alta */
    h1, h2, h3, h4 {
        color: #FFFFFF !important;
        font-weight: 900 !important;
        letter-spacing: 0.5px;
    }
    
    /* Botões principais no estilo amarelo ouro com texto preto */
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
    
    /* Estilização dos acordeões (Expanders) */
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

    /* Inputs de texto e selects */
    .stTextInput input, .stSelectbox [data-baseweb="select"] {
        background-color: #1A1A1A !important;
        color: #FFFFFF !important;
        border: 1px solid #333333 !important;
        border-radius: 4px !important;
    }
    
    /* Multiselect tags */
    span[data-baseweb="tag"] {
        background-color: #FDD818 !important;
        color: #000000 !important;
        font-weight: bold !important;
    }

    /* Cards de métricas */
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
# DADOS DOS PROMOTORES
# ==============================================================================
DADOS_PROMOTORES = {
    "Pamela Camila de Almeida Alexandrino": {
        "endereco": "Rua Doutor Rowilson Flora, 753 - Poços de Caldas/MG",
        "lat": -21.7858,
        "lon": -46.5625,
        "cidades": [
            "POÇOS DE CALDAS", "ANDRADAS", "GUAXUPÉ", "ITAJUBA", "POUSO ALEGRE",
            "VARGINHA", "TRÊS PONTAS", "TRÊS CORAÇÕES", "MACHADO", "ALFENAS"
        ]
    },
    "Fernanda Dias Ferreira": {
        "endereco": "Rua Jorge Raimundo, 409 - Juiz de Fora/MG",
        "lat": -21.7642,
        "lon": -43.3496,
        "cidades": [
            "JUIZ DE FORA"
        ]
    },
    "Saruete Valeska Stabile de Oliveira": {
        "endereco": "Rua José Gonçalves de Souza, 105 - São José do Rio Preto/SP",
        "lat": -20.8113,
        "lon": -49.3758,
        "cidades": [
            "SÃO JOSÉ DO RIO PRETO", "MIRASSOL", "CATANDUVA"
        ]
    },
    "Carolina Rodrigues Bruno": {
        "endereco": "Rua Doutor Bernardino de Campos - São Carlos/SP",
        "lat": -22.0175,
        "lon": -47.8908,
        "cidades": [
            "SÃO CARLOS", "ARARAQUARA", "MATÃO"
        ]
    },
    "Madalla Teixeira Reis": {
        "endereco": "Rua Odilon Machado, 105 - Tocantins/MG",
        "lat": -21.1764,
        "lon": -43.0181,
        "cidades": [
            "UBÁ", "DESCOBERTO", "SÃO JOÃO NEPOMUCENO", "VIÇOSA",
            "TOCANTINS", "RODEIRO", "PIRAÚBA", "GUARANI", "RIO POMBA", "RIO NOVO"
        ]
    },
    "Rodrigo Luis Adao": {
        "endereco": "Avenida Professora Edul Rangel Rabello, 405 - Ribeirão Preto/SP",
        "lat": -21.2075,
        "lon": -47.7981,
        "cidades": [
            "RIBEIRÃO PRETO"
        ]
    }
}

PROMOTORES = list(DADOS_PROMOTORES.keys())
SITUACOES = ['Normal', 'Férias', 'Carro Quebrado', 'Feriado', 'Atestado Médico', 'Folga', 'Falta']
NOME_PLANILHA_CLIENTES = "Cópia de clientes com cnpj corretinho novinho (1).xlsx"

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
# TELA DE IDENTIFICAÇÃO (QUEM É VOCÊ?)
# ==============================================================================
if "usuario_ativo" not in st.session_state:
    st.session_state.usuario_ativo = None

if not st.session_state.usuario_ativo:
    st.markdown("<h1 style='text-align: center; color: #FDD818 !important;'># ACESSO DE PROMOTORES</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #888;'>SELECIONE SEU PERFIL PARA INICIAR O REGISTRO</p>", unsafe_allow_html=True)
    
    escolha_promotor = st.selectbox(
        "QUEM É VOCÊ?",
        options=["-- Selecione seu nome --"] + PROMOTORES
    )
    
    if st.button("ACESSAR SISTEMA ➔", type="primary"):
        if escolha_promotor != "-- Selecione seu nome --":
            st.session_state.usuario_ativo = escolha_promotor
            st.rerun()
        else:
            st.warning("Por favor, selecione seu nome na lista antes de prosseguir.")
    st.stop()

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

usuario_logado = st.session_state.usuario_ativo
dados_promotor_atual = DADOS_PROMOTORES.get(usuario_logado, {})
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

# ==============================================================================
# INTEGRAÇÃO COM GITHUB
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

def carregar_dados_github(caminho_arquivo):
    token, repo, branch = get_github_credentials()
    if not token or not repo:
        return None, None

    url = f"https://api.github.com/repos/{repo}/contents/{caminho_arquivo}?ref={branch}"
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
    return None, None

def salvar_dados_github(caminho_arquivo, dados_dict, sha_existente=None, mensagem_commit="Atualização de KM"):
    token, repo, branch = get_github_credentials()
    if not token or not repo:
        st.error("Credenciais do GitHub não configuradas nos Secrets do Streamlit!")
        return False

    url = f"https://api.github.com/repos/{repo}/contents/{caminho_arquivo}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github.v3+json"}

    conteudo_json_str = json.dumps(dados_dict, indent=4, ensure_ascii=False)
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

# ==============================================================================
# CABEÇALHO DO APLICATIVO
# ==============================================================================
promotor_sel = st.session_state.usuario_ativo

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

nome_prom_limpo = promotor_sel.replace(" ", "_")
caminho_github = f"dados_promotores/S{num_semana}_{nome_prom_limpo}.json"

dados_salvos, sha_arquivo = carregar_dados_github(caminho_github)

esta_finalizado = False
if dados_salvos and dados_salvos.get("status") == "FINALIZADO":
    esta_finalizado = True
    st.success("🔒 **SEMANA FINALIZADA E TRANSMITIDA.** Todos os campos estão bloqueados.")
    if st.button("🔓 REABRIR PARA CORREÇÃO"):
        sucesso_reabrir = salvar_dados_github(
            caminho_github,
            {**dados_salvos, "status": "RASCUNHO"},
            sha_existente=sha_arquivo,
            mensagem_commit=f"Reaberto para edição S{num_semana} - {promotor_sel}"
        )
        if sucesso_reabrir:
            st.success("Semana destravada com sucesso!")
            st.rerun()
elif dados_salvos:
    st.warning("📝 Rascunho salvo em aberto. Edição liberada.")

st.divider()

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
                    st.markdown(f"• **{row['NOME']}**  \n  🏠 {row['ENDEREÇO']} - {row['BAIRRO']}, {row['CIDADE_RAW']}/{row['UF']}")

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
                    get_color=[253, 216, 24, 230],  # Amarelo ouro nas paradas
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
                    get_color=[255, 255, 255, 160],  # Linhas brancas de alta visibilidade
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
# GASTOS EXTRAS
# ==============================================================================
st.divider()
st.markdown("### 💰 GASTOS EXTRAS")

gastos_salvos_default = dados_salvos.get("gastos_extras", []) if dados_salvos else []
qtd_gastos = max(1, len(gastos_salvos_default))

gastos_extras = []
for idx in range(qtd_gastos):
    g_item = gastos_salvos_default[idx] if idx < len(gastos_salvos_default) else {}
    col_desc, col_val = st.columns([3, 2])
    
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

# ==============================================================================
# RESUMO FINANCEIRO
# ==============================================================================
st.divider()
st.markdown("### 📊 FECHAMENTO DA SEMANA")

VALOR_KM_TAXA = 1.17
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
            sucesso = salvar_dados_github(
                caminho_github, 
                payload, 
                sha_existente=sha_arquivo,
                mensagem_commit=f"Rascunho S{num_semana} - {promotor_sel}"
            )
            if sucesso:
                st.success("Rascunho salvo com sucesso!")
                st.rerun()

    with col_btn2:
        if st.button("FINALIZAR SEMANA 🚀", type="primary"):
            payload = construir_payload("FINALIZADO")
            sucesso = salvar_dados_github(
                caminho_github, 
                payload, 
                sha_existente=sha_arquivo,
                mensagem_commit=f"FINALIZADO S{num_semana} - {promotor_sel}"
            )
            if sucesso:
                st.balloons()
                st.success("Semana finalizada e bloqueada com sucesso!")
                st.rerun()
else:
    st.info("ℹ️ Para realizar qualquer edição, clique em '🔓 REABRIR PARA CORREÇÃO' no topo da tela.")
