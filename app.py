import streamlit as st
import requests
import json
import base64
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

# Estilo visual limpo para dispositivos móveis
st.markdown("""
    <style>
    .main { padding: 10px; }
    .stButton>button { width: 100%; border-radius: 8px; font-weight: bold; }
    .stTextInput input { text-align: center; }
    </style>
""", unsafe_allow_html=True)

# Lista de Promotores
PROMOTORES = [
    "Fernanda Dias Ferreira",
    "João Silva",
    "Maria Souza",
    "Carlos Oliveira"
]

SITUACOES = ['Normal', 'Férias', 'Carro Quebrado', 'Feriado', 'Atestado Médico', 'Folga', 'Falta']

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
    
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json()
        conteudo_base64 = data.get("content", "")
        sha = data.get("sha", "")
        conteudo_json = json.loads(base64.b64decode(conteudo_base64).decode('utf-8'))
        return conteudo_json, sha
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

    response = requests.put(url, headers=headers, json=payload)
    return response.status_code in [200, 201]

# ==============================================================================
# AUXILIARES
# ==============================================================================
def calcular_intervalo_semana(num_semana, ano=2026):
    start = datetime(ano, 1, 1)
    start -= timedelta(days=start.weekday())
    segunda = start + timedelta(weeks=num_semana - 1)
    domingo = segunda + timedelta(days=6)
    return segunda, domingo

# ==============================================================================
# INTERFACE PRINCIPAL
# ==============================================================================
st.title("🚗 Controle de KM")
st.subheader("Minassal - Lançamento Semanal")

col1, col2 = st.columns([2, 1])

with col1:
    promotor_sel = st.selectbox("Promotor(a):", PROMOTORES)

semana_atual_default = datetime.now().isocalendar()[1]
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

mapa_dados_salvos = {}
if dados_salvos and "detalhes" in dados_salvos:
    mapa_dados_salvos = {d["dia"]: d for d in dados_salvos["detalhes"]}

st.markdown("### 📋 Registros Diários")

km_fim_anterior = "0"

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
            val_kmi_default = dados_dia_salvo.get("km_ini", km_fim_anterior)
            km_ini_str = st.text_input(f"KM Inicial", value=str(val_kmi_default), key=f"kmi_{dia_nome}")

        with col_kmf:
            val_kmf_default = dados_dia_salvo.get("km_fim", "0")
            km_fim_str = st.text_input(f"KM Final", value=str(val_kmf_default), key=f"kmf_{dia_nome}")

        if km_fim_str and km_fim_str != "0":
            km_fim_anterior = km_fim_str

        cli_default = ", ".join(dados_dia_salvo.get("clientes", [])) if isinstance(dados_dia_salvo.get("clientes"), list) else ""
        clientes_str = st.text_input("Códigos dos Clientes:", value=cli_default, key=f"cli_{dia_nome}")

        try:
            k_i = float(km_ini_str.replace(',', '.'))
            k_f = float(km_fim_str.replace(',', '.'))
            km_dia = max(0.0, k_f - k_i)
        except ValueError:
            km_dia = 0.0

        if km_dia > 0:
            st.caption(f"🚘 KM Rodado no dia: **{km_dia:.1f} km**")

        km_total_calculado += km_dia

        lista_cods = [c.strip() for c in clientes_str.split(',') if c.strip()]
        detalhes_dias.append({
            "dia": dia_nome,
            "data": data_dia,
            "km": km_dia,
            "sit": sit_sel,
            "clientes": lista_cods,
            "leitura": lei_sel,
            "km_ini": km_ini_str,
            "km_fim": km_fim_str
        })

# GASTOS EXTRAS
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
        g_val = st.text_input(f"Valor R$ #{idx+1}", value=str(g_item.get("valor", "0.00")), key=f"gval_{idx}")
    with col_inf:
        g_inf = st.checkbox("Info", value=g_item.get("informativo", False), key=f"ginf_{idx}")

    if g_desc.strip():
        try:
            v_float = float(g_val.replace(',', '.'))
        except ValueError:
            v_float = 0.0
        gastos_extras.append({"desc": g_desc.strip(), "valor": v_float, "informativo": g_inf})

# RESUMO DE VALORES
st.divider()
VALOR_KM_TAXA = 1.17
valor_total_km = km_total_calculado * VALOR_KM_TAXA
valor_extras_reembolsavel = sum(g["valor"] for g in gastos_extras if not g.get("informativo", False))
valor_total_geral = valor_total_km + valor_extras_reembolsavel

st.metric(label="Total KM Rodado na Semana", value=f"{km_total_calculado:.1f} km")
st.metric(label="Valor Total de Reembolso Estimado", value=f"R$ {valor_total_geral:.2f}")

# BOTÕES DE AÇÃO
col_btn1, col_btn2 = st.columns(2)

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
            st.success("Rascunho salvo no GitHub com sucesso!")
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
            st.success("Semana FINALIZADA e enviada com sucesso!")
            st.rerun()
