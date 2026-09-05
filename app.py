import streamlit as st
import pandas as pd
import requests
import json
import base64
import glob
from datetime import datetime, timedelta

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA (OTIMIZADA PARA CELULAR)
# ==============================================================================
st.set_page_config(
    page_title="Minassal - Controle de KM",
    page_icon="🚗",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    .main { padding: 10px; }
    .stButton>button { width: 100%; border-radius: 8px; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# CONFIGURAÇÕES E CARREGAMENTO DINÂMICO DA PLANILHA
# ==============================================================================
PROMOTORES = [
    "Fernanda Dias Ferreira",
    "João Silva",
    "Maria Souza",
    "Carlos Oliveira"
]

SITUACOES = ['Normal', 'Férias', 'Carro Quebrado', 'Feriado', 'Atestado Médico', 'Folga', 'Falta']
NOME_ARQUIVO_PLANILHA = "Cópia de clientes com cnpj corretinho novinho (1).xlsx"

@st.cache_data(ttl=3600)
def carregar_base_clientes():
    caminho = NOME_ARQUIVO_PLANILHA
    arquivos_excel = glob.glob("*.xlsx")
    
    # Se o nome exato não bater (devido a encoding do Linux/GitHub), busca qualquer arquivo de clientes
    if caminho not in arquivos_excel and arquivos_excel:
        for arq in arquivos_excel:
            if "clientes" in arq.lower():
                caminho = arq
                break
        else:
            caminho = arquivos_excel[0]

    try:
        df = pd.read_excel(caminho, sheet_name=0)
        df = df.dropna(subset=["CÓDIGO", "NOME"]).copy()
        
        # Limpeza e padronização dos campos
        df["CÓDIGO"] = df["CÓDIGO"].astype(int).astype(str).str.strip()
        df["NOME"] = df["NOME"].astype(str).str.strip()
        df["CIDADE"] = df["CIDADE"].fillna("NÃO INFORMADA").astype(str).str.strip().str.upper()
        df["BAIRRO"] = df["BAIRRO"].fillna("").astype(str).str.strip()
        df["ENDEREÇO"] = df["ENDEREÇO"].fillna("").astype(str).str.strip()
        df["UF"] = df["UF"].fillna("").astype(str).str.strip()

        # Conversão de coordenadas (vírgula para ponto)
        def sanitizar_coord(val):
            try:
                if pd.isna(val):
                    return None
                return float(str(val).replace(",", ".").strip())
            except ValueError:
                return None

        df["lat"] = df["LATITUDE"].apply(sanitizar_coord)
        df["lon"] = df["LONGITUDE"].apply(sanitizar_coord)
        return df
    except Exception as e:
        st.error(f"Erro ao carregar a planilha '{caminho}': {e}")
        return pd.DataFrame()

DF_CLIENTES = carregar_base_clientes()
LISTA_CIDADES = sorted([c for c in DF_CLIENTES["CIDADE"].unique() if c]) if not DF_CLIENTES.empty else []

MAPA_GERAL_NOMES = {}
if not DF_CLIENTES.empty:
    for _, r in DF_CLIENTES.iterrows():
        bairro_str = f" - {r['BAIRRO']}" if r['BAIRRO'] else ""
        cidade_str = f" ({r['CIDADE']}/{r['UF']})" if r['CIDADE'] else ""
        MAPA_GERAL_NOMES[r["CÓDIGO"]] = f"{r['NOME']}{bairro_str}{cidade_str}"

# ==============================================================================
# INTEGRAÇÃO COM A API DO GITHUB
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
# AUXILIARES
# ==============================================================================
def calcular_intervalo_semana(num_semana, ano=None):
    if ano is None:
        ano = datetime.now().year
    start = datetime(ano, 1, 1)
    start -= timedelta(days=start.weekday())
    segunda = start + timedelta(weeks=num_semana - 1)
    domingo = segunda + timedelta(days=6)
    return segunda, domingo

def converter_float(val):
    try:
        return float(str(val).replace(',', '.').strip())
    except (ValueError, TypeError):
        return 0.0

# ==============================================================================
# INTERFACE PRINCIPAL
# ==============================================================================
st.title("🚗 Controle de KM")
st.subheader("Minassal - Lançamento Semanal")

col1, col2 = st.columns([2, 1])
with col1:
    promotor_sel = st.selectbox("Promotor(a):", PROMOTORES)

semana_atual_default = int(datetime.now().isocalendar()[1])
with col2:
    num_semana = st.number_input("Nº Semana:", min_value=1, max_value=53, value=semana_atual_default)

segunda, domingo = calcular_intervalo_semana(num_semana)
intervalo_str = f"{segunda.strftime('%d/%m')} a {domingo.strftime('%d/%m')}"
st.info(f"📅 **Período:** Semana {num_semana} ({intervalo_str})")

nome_prom_limpo = promotor_sel.replace(" ", "_")
caminho_github = f"dados_promotores/S{num_semana}_{nome_prom_limpo}.json"

dados_salvos, sha_arquivo = carregar_dados_github(caminho_github)

if dados_salvos:
    if dados_salvos.get("status") == "FINALIZADO":
        st.success("✅ Esta semana já foi FINALIZADA pelo promotor.")
    else:
        st.warning("📝 Rascunho salvo anteriormente carregado!")

st.divider()

dias_semana = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
detalhes_dias = []
km_total_calculado = 0.0
km_fim_anterior = 0.0

mapa_dados_salvos = {}
if dados_salvos and "detalhes" in dados_salvos:
    mapa_dados_salvos = {d["dia"]: d for d in dados_salvos["detalhes"]}

st.markdown("### 📋 Registros Diários")

for i, dia_nome in enumerate(dias_semana):
    data_dia = (segunda + timedelta(days=i)).strftime("%d/%m")
    dados_dia_salvo = mapa_dados_salvos.get(dia_nome, {})

    with st.expander(f"📌 **{dia_nome} ({data_dia})**", expanded=(i == 0 or bool(dados_dia_salvo))):
        col_sit, col_lei = st.columns([2, 1])
        with col_sit:
            sit_default = dados_dia_salvo.get("sit", "Normal")
            idx_sit = SITUACOES.index(sit_default) if sit_default in SITUACOES else 0
            sit_sel = st.selectbox(f"Situação ({dia_nome})", SITUACOES, index=idx_sit, key=f"sit_{dia_nome}")

        with col_lei:
            lei_default = dados_dia_salvo.get("leitura", False)
            lei_sel = st.checkbox("Leituras?", value=lei_default, key=f"lei_{dia_nome}")

        col_kmi, col_kmf = st.columns(2)
        with col_kmi:
            def_kmi = converter_float(dados_dia_salvo.get("km_ini", km_fim_anterior))
            km_ini = st.number_input(f"KM Inicial ({dia_nome})", min_value=0.0, value=def_kmi, step=1.0, key=f"kmi_{dia_nome}")

        with col_kmf:
            def_kmf = converter_float(dados_dia_salvo.get("km_fim", def_kmi))
            km_fim = st.number_input(f"KM Final ({dia_nome})", min_value=0.0, value=def_kmf, step=1.0, key=f"kmf_{dia_nome}")

        km_dia = 0.0
        if km_fim > 0.0:
            if km_fim < km_ini:
                st.error("⚠️ KM Final não pode ser menor que o KM Inicial!")
            else:
                km_dia = km_fim - km_ini
                km_fim_anterior = km_fim
                st.caption(f"🚘 KM Rodado no dia: **{km_dia:.1f} km**")

        km_total_calculado += km_dia

        # --- SELEÇÃO DE CLIENTES FILTRADA POR CIDADE ---
        cidade_sel = st.selectbox(
            f"Filtrar por Cidade ({dia_nome}):",
            options=["TODAS"] + LISTA_CIDADES,
            key=f"cid_{dia_nome}"
        )

        if cidade_sel != "TODAS" and not DF_CLIENTES.empty:
            df_opcoes = DF_CLIENTES[DF_CLIENTES["CIDADE"] == cidade_sel]
        else:
            df_opcoes = DF_CLIENTES

        opcoes_cods = df_opcoes["CÓDIGO"].tolist() if not df_opcoes.empty else []
        cods_salvos = [str(c) for c in dados_dia_salvo.get("clientes", [])]
        opcoes_finais = list(dict.fromkeys(cods_salvos + opcoes_cods))

        clientes_sel = st.multiselect(
            f"Lojas Atendidas ({dia_nome}):",
            options=opcoes_finais,
            default=cods_salvos,
            format_func=lambda cod: MAPA_GERAL_NOMES.get(cod, f"Cód: {cod}"),
            key=f"cli_{dia_nome}"
        )

        # --- EXIBIÇÃO DE ENDEREÇOS E MAPA DO DIA ---
        if clientes_sel and not DF_CLIENTES.empty:
            df_atendidos = DF_CLIENTES[DF_CLIENTES["CÓDIGO"].isin(clientes_sel)]

            with st.expander(f"📍 Endereços Visitados ({len(clientes_sel)})"):
                for _, row in df_atendidos.iterrows():
                    st.markdown(f"**{row['NOME']}**  \n🏠 {row['ENDEREÇO']} - {row['BAIRRO']}, {row['CIDADE']}/{row['UF']}")

            df_mapa = df_atendidos.dropna(subset=["lat", "lon"])
            if not df_mapa.empty:
                st.caption("🗺️ Rota / Locais no Mapa:")
                st.map(df_mapa[["lat", "lon"]], zoom=11)

        detalhes_dias.append({
            "dia": dia_nome,
            "data": data_dia,
            "km": km_dia,
            "sit": sit_sel,
            "clientes": clientes_sel,
            "leitura": lei_sel,
            "km_ini": km_ini,
            "km_fim": km_fim
        })

# ==============================================================================
# GASTOS EXTRAS
# ==============================================================================
st.divider()
st.markdown("### 💰 Gastos Extras (Opção)")

gastos_salvos_default = dados_salvos.get("gastos_extras", []) if dados_salvos else []
qtd_gastos = max(1, len(gastos_salvos_default))

gastos_extras = []
for idx in range(qtd_gastos):
    g_item = gastos_salvos_default[idx] if idx < len(gastos_salvos_default) else {}
    col_desc, col_val, col_inf = st.columns([3, 2, 1])
    
    with col_desc:
        g_desc = st.text_input(f"Descrição #{idx+1}", value=g_item.get("desc", ""), key=f"gdesc_{idx}")
    with col_val:
        v_salvo = converter_float(g_item.get("valor", 0.0))
        g_val = st.number_input(f"Valor R$ #{idx+1}", min_value=0.0, value=v_salvo, step=1.0, key=f"gval_{idx}")
    with col_inf:
        g_inf = st.checkbox("Info", value=g_item.get("informativo", False), key=f"ginf_{idx}")

    if g_desc.strip():
        gastos_extras.append({"desc": g_desc.strip(), "valor": g_val, "informativo": g_inf})

# ==============================================================================
# RESUMO E AÇÕES
# ==============================================================================
st.divider()
VALOR_KM_TAXA = 1.17
valor_total_km = km_total_calculado * VALOR_KM_TAXA
valor_extras_reembolsavel = sum(g["valor"] for g in gastos_extras if not g.get("informativo", False))
valor_total_geral = valor_total_km + valor_extras_reembolsavel

col_m1, col_m2 = st.columns(2)
col_m1.metric("Total KM", f"{km_total_calculado:.1f} km")
col_m2.metric("Total Reembolso", f"R$ {valor_total_geral:.2f}")

def construir_payload(status_envio):
    return {
        "id": int(datetime.now().timestamp()),
        "semana_ref": str(num_semana),
        "intervalo_datas": intervalo_str,
        "promotor": promotor_sel,
        "status": status_envio,
        "km_total": km_total_calculado,
        "valor_total": valor_total_geral,
        "gastos_extras": gastos_extras,
        "detalhes": detalhes_dias
    }

col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    if st.button("💾 Salvar Rascunho"):
        payload = construir_payload("RASCUNHO")
        sucesso = salvar_dados_github(
            caminho_github, 
            payload, 
            sha_existente=sha_arquivo,
            mensagem_commit=f"Rascunho S{num_semana} - {promotor_sel}"
        )
        if sucesso:
            st.success("Rascunho salvo no GitHub!")
            st.rerun()

with col_btn2:
    if st.button("🚀 Finalizar Semana", type="primary"):
        payload = construir_payload("FINALIZADO")
        sucesso = salvar_dados_github(
            caminho_github, 
            payload, 
            sha_existente=sha_arquivo,
            mensagem_commit=f"FINALIZADO S{num_semana} - {promotor_sel}"
        )
        if sucesso:
            st.balloons()
            st.success("Semana finalizada com sucesso!")
            st.rerun()
