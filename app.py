import streamlit as st
import pandas as pd
import requests
import json
import base64
import glob
import unicodedata
import pydeck as pdk
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
    div[data-testid="stExpander"] div[role="button"] p { font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# DADOS DOS PROMOTORES (ENDEREÇOS E CIDADES DA ROTA)
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
NOME_PLANILHA_VENDAS = "cubo_de_vendas_05_09_2026_13_31_05.xlsx"

# ==============================================================================
# AUXILIARES DE FORMATAÇÃO
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

# ==============================================================================
# TELA DE IDENTIFICAÇÃO (QUEM É VOCÊ?)
# ==============================================================================
if "usuario_ativo" not in st.session_state:
    st.session_state.usuario_ativo = None

if not st.session_state.usuario_ativo:
    st.markdown("### 👤 Identificação")
    st.markdown("#### QUEM É VOCÊ?")
    
    escolha_promotor = st.selectbox(
        "Selecione o seu nome na lista para continuar:",
        options=["-- Selecione seu nome --"] + PROMOTORES
    )
    
    if st.button("Acessar Sistema 🚀", type="primary"):
        if escolha_promotor != "-- Selecione seu nome --":
            st.session_state.usuario_ativo = escolha_promotor
            st.rerun()
        else:
            st.warning("Por favor, selecione seu nome na lista antes de prosseguir.")
    st.stop()

# ==============================================================================
# CARREGAMENTO DA BASE CRUZADA (FILTRADA E ORDENADA POR COMPRAS, SEM VALOR NA TELA)
# ==============================================================================
@st.cache_data(ttl=3600)
def carregar_base_cruzada():
    caminho_cli = NOME_PLANILHA_CLIENTES
    arqs = glob.glob("*.xlsx")
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

    # Leitura do Cubo de Vendas para filtrar quem comprou e ordenar
    caminho_vendas = NOME_PLANILHA_VENDAS
    if caminho_vendas not in arqs and arqs:
        for a in arqs:
            if "cubo" in a.lower() or "vendas" in a.lower():
                caminho_vendas = a
                break

    try:
        df_vendas = pd.read_excel(caminho_vendas, sheet_name=0)
        df_vendas = df_vendas.dropna(subset=["CLIENTE CODIGO", "TOTAL VALOR"]).copy()
        df_vendas["CLIENTE CODIGO"] = df_vendas["CLIENTE CODIGO"].astype(int).astype(str).str.strip()
        
        # Filtra apenas clientes com valor positivo no ano
        df_vendas = df_vendas[df_vendas["TOTAL VALOR"] > 0]
        vendas_resumo = df_vendas.groupby("CLIENTE CODIGO")["TOTAL VALOR"].sum().reset_index()
    except Exception:
        return df_cli

    # Junta e ordena por maior valor de compras sem exibir o valor ao usuário
    df_merged = pd.merge(df_cli, vendas_resumo, left_on="CÓDIGO", right_on="CLIENTE CODIGO", how="inner")
    df_merged = df_merged.sort_values(by="TOTAL VALOR", ascending=False)
    return df_merged

DF_CLIENTES = carregar_base_cruzada()

# Rótulo limpo: Apenas Nome, Bairro e Cidade (sem valores de compra)
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
    st.title("🚗 Controle de KM")
with col_sair:
    st.write("")
    if st.button("Trocar Promotor 🔄"):
        st.session_state.usuario_ativo = None
        st.rerun()

st.caption(f"👤 Promotor(a): **{promotor_sel}**")
st.caption(f"🏠 Residência: {dados_promotor_atual.get('endereco', 'Não cadastrado')}")

semana_atual_default = int(datetime.now().isocalendar()[1])
num_semana = st.number_input("Nº da Semana:", min_value=1, max_value=53, value=semana_atual_default)

segunda, domingo = calcular_intervalo_semana(num_semana)
intervalo_str = f"{segunda.strftime('%d/%m')} a {domingo.strftime('%d/%m')}"
st.info(f"📅 **Período:** Semana {num_semana} ({intervalo_str})")

nome_prom_limpo = promotor_sel.replace(" ", "_")
caminho_github = f"dados_promotores/S{num_semana}_{nome_prom_limpo}.json"

dados_salvos, sha_arquivo = carregar_dados_github(caminho_github)

if dados_salvos:
    if dados_salvos.get("status") == "FINALIZADO":
        st.success("✅ Esta semana já foi FINALIZADA.")
    else:
        st.warning("📝 Rascunho salvo anteriormente carregado!")

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
            def_kmi = str_br_para_float(dados_dia_salvo.get("km_ini", km_fim_anterior))
            km_ini_str = st.text_input(f"KM Inicial ({dia_nome})", value=float_para_str_br(def_kmi).replace(",00", ""), key=f"kmi_{dia_nome}")
            km_ini = str_br_para_float(km_ini_str)

        with col_kmf:
            def_kmf = str_br_para_float(dados_dia_salvo.get("km_fim", def_kmi))
            km_fim_str = st.text_input(f"KM Final ({dia_nome})", value=float_para_str_br(def_kmf).replace(",00", ""), key=f"kmf_{dia_nome}")
            km_fim = str_br_para_float(km_fim_str)

        km_dia = 0.0
        if km_fim > 0.0:
            if km_fim < km_ini:
                st.error("⚠️ KM Final não pode ser menor que o KM Inicial!")
            else:
                km_dia = km_fim - km_ini
                km_fim_anterior = km_fim
                st.caption(f"🚘 KM Rodado no dia: **{float_para_str_br(km_dia)} km**")

        km_total_calculado += km_dia

        # Submenu por Cidade (Lojas compradoras ordenadas internamente por volume)
        st.markdown("#### 🏬 Selecionar Lojas por Cidade")
        cods_salvos_dia = [str(c) for c in dados_dia_salvo.get("clientes", [])]
        clientes_dia_selecionados = []

        if not DF_CLIENTES.empty:
            for cid_norm in cidades_disponiveis_promotor:
                df_cidade = DF_CLIENTES[DF_CLIENTES["CIDADE_NORM"] == cid_norm]
                if df_cidade.empty:
                    continue
                
                nome_cidade_exibicao = df_cidade["CIDADE_RAW"].iloc[0]
                # A lista de códigos já está ordenada pelo volume de compras
                codigos_da_cidade = df_cidade["CÓDIGO"].tolist()
                defaults_cidade = [c for c in cods_salvos_dia if c in codigos_da_cidade]

                with st.expander(f"🏙️ {nome_cidade_exibicao} ({len(df_cidade)} lojas)", expanded=bool(defaults_cidade)):
                    escolhidos_cid = st.multiselect(
                        f"Lojas em {nome_cidade_exibicao}:",
                        options=codigos_da_cidade,
                        default=defaults_cidade,
                        format_func=lambda cod: MAPA_GERAL_NOMES.get(cod, f"Cód: {cod}"),
                        key=f"cli_{dia_nome}_{cid_norm}"
                    )
                    clientes_dia_selecionados.extend(escolhidos_cid)

            # Cidade fora da rota usual
            with st.expander("🌐 Outra Cidade / Fora da Rota Usual", expanded=False):
                cidades_fora = [c for c in todas_cidades_norm if c not in cidades_disponiveis_promotor]
                cid_extra = st.selectbox(f"Escolha a cidade fora da rota ({dia_nome}):", ["-- Selecione --"] + cidades_fora, key=f"extra_cid_{dia_nome}")
                if cid_extra != "-- Selecione --":
                    df_extra = DF_CLIENTES[DF_CLIENTES["CIDADE_NORM"] == cid_extra]
                    cods_extra = df_extra["CÓDIGO"].tolist()
                    defaults_extra = [c for c in cods_salvos_dia if c in cods_extra]
                    escolhidos_extra = st.multiselect(
                        f"Lojas em {cid_extra}:",
                        options=cods_extra,
                        default=defaults_extra,
                        format_func=lambda cod: MAPA_GERAL_NOMES.get(cod, f"Cód: {cod}"),
                        key=f"cli_extra_{dia_nome}"
                    )
                    clientes_dia_selecionados.extend(escolhidos_extra)

        clientes_dia_selecionados = list(dict.fromkeys(clientes_dia_selecionados))

        # Mapa: Linha ligando a residência às lojas atendidas
        if clientes_dia_selecionados and not DF_CLIENTES.empty:
            df_atendidos = DF_CLIENTES[DF_CLIENTES["CÓDIGO"].isin(clientes_dia_selecionados)]

            with st.expander(f"📍 Endereços Selecionados ({len(clientes_dia_selecionados)})", expanded=True):
                for _, row in df_atendidos.iterrows():
                    st.markdown(f"**{row['NOME']}**  \n🏠 {row['ENDEREÇO']} - {row['BAIRRO']}, {row['CIDADE_RAW']}/{row['UF']}")

            df_coords = df_atendidos.dropna(subset=["lat", "lon"]).copy()

            if not df_coords.empty and "lat" in dados_promotor_atual:
                st.caption("🗺️ Rota: Casa do Promotor (Azul) ➔ Lojas Atendidas (Vermelho):")
                lat_casa = dados_promotor_atual["lat"]
                lon_casa = dados_promotor_atual["lon"]

                camada_casa = pdk.Layer(
                    "ScatterplotLayer",
                    data=pd.DataFrame([{"lat": lat_casa, "lon": lon_casa}]),
                    get_position=["lon", "lat"],
                    get_color=[0, 100, 255, 200],
                    get_radius=800,
                    pickable=True
                )

                camada_lojas = pdk.Layer(
                    "ScatterplotLayer",
                    data=df_coords,
                    get_position=["lon", "lat"],
                    get_color=[230, 40, 40, 200],
                    get_radius=500,
                    pickable=True
                )

                linhas_rotas = [{"origem": [lon_casa, lat_casa], "destino": [r["lon"], r["lat"]]} for _, r in df_coords.iterrows()]
                camada_linhas = pdk.Layer(
                    "LineLayer",
                    data=pd.DataFrame(linhas_rotas),
                    get_source_position="origem",
                    get_target_position="destino",
                    get_color=[30, 30, 30, 160],
                    get_width=3
                )

                viewport = pdk.ViewState(latitude=lat_casa, longitude=lon_casa, zoom=10, pitch=0)
                st.pydeck_chart(pdk.Deck(layers=[camada_linhas, camada_casa, camada_lojas], initial_view_state=viewport, map_style="road"))

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
# GASTOS EXTRAS (DIGITAÇÃO COM VÍRGULA PADRÃO BRASIL)
# ==============================================================================
st.divider()
st.markdown("### 💰 Gastos Extras da Semana")
st.caption("Digite o valor com **vírgula** (Exemplo: `25,50` ou `150,00`).")

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
            key=f"gdesc_{idx}"
        )
    with col_val:
        v_salvo_num = g_item.get("valor", 0.0)
        v_salvo_txt = float_para_str_br(v_salvo_num) if v_salvo_num > 0 else ""
        
        g_val_txt = st.text_input(
            f"Valor R$ #{idx+1}", 
            value=v_salvo_txt, 
            placeholder="0,00", 
            key=f"gval_{idx}"
        )
        v_float = str_br_para_float(g_val_txt)

    if g_desc.strip() and v_float > 0:
        gastos_extras.append({"desc": g_desc.strip(), "valor": v_float})

# ==============================================================================
# RESUMO FINANCEIRO (COM VÍRGULA)
# ==============================================================================
st.divider()
st.markdown("### 📊 Fechamento da Semana")

VALOR_KM_TAXA = 1.17
valor_total_km = km_total_calculado * VALOR_KM_TAXA
valor_extras_total = sum(g["valor"] for g in gastos_extras)
valor_total_reembolso = valor_total_km + valor_extras_total

c_km, c_ext, c_tot = st.columns(3)
c_km.metric("Reembolso KM", f"R$ {float_para_str_br(valor_total_km)}", help=f"{float_para_str_br(km_total_calculado)} km x R$ 1,17")
c_ext.metric("Gastos Extras", f"R$ {float_para_str_br(valor_extras_total)}")
c_tot.metric("Total a Receber", f"R$ {float_para_str_br(valor_total_reembolso)}")

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
