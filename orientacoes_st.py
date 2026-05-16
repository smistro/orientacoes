import streamlit as st
import pandas as pd
from datetime import date, datetime
import plotly.express as px
import gspread
from google.oauth2.service_account import Credentials

# =========================================================
# SISTEMA COLABORATIVO DE CONTROLE DE ORIENTAÇÕES
# Streamlit + Google Sheets
# Versão com autenticação interna por e-mail e senha
# =========================================================

st.set_page_config(
    page_title="Controle de Orientações",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# CONFIGURAÇÕES
# =========================================================

# Nome da planilha Google exatamente como aparece no Google Drive
NOME_PLANILHA = "controle_orientacoes"

# Nomes das abas obrigatórias na planilha
ABA_CADASTRO = "cadastro"
ABA_REGISTROS = "registros"
ABA_CONFIG = "configuracoes"

# A aba cadastro precisa ter exatamente estas colunas na primeira linha:
# Discente | Programa | Nível | Email | Senha | Perfil | Ativo
COLUNAS_CADASTRO = [
    "Discente",
    "Programa",
    "Nível",
    "Email",
    "Senha",
    "Perfil",
    "Ativo"
]

COLUNAS_REGISTROS = [
    "Data", "Hora", "Discente", "Programa", "Nível", "Email",
    "Situação", "Pendências", "Responsável", "Prazo",
    "Prioridade", "Observações", "Atualizado_por"
]

novo_registro = {
    "Data": data_hoje,
    "Hora": hora_agora,
    "Discente": discente,
    "Programa": programa,
    "Nível": nivel,
    "Email": email,
    "Situação": situacao,
    "Pendências": pendencias,
    "Responsável": responsavel,
    "Prazo": prazo,
    "Prioridade": prioridade,
    "Observações": observacoes,
    "Atualizado_por": email_usuario
}

df_novo = pd.DataFrame([novo_registro], columns=COLUNAS_REGISTROS)

PERFIS = ["Orientando", "Orientador"]

SITUACOES_PADRAO = [
    "Projeto de pesquisa",
    "Submissão ao CEP",
    "Coleta de dados",
    "Análise dos dados",
    "Escrita do artigo",
    "Escrita da dissertação/TCC",
    "Qualificação",
    "Defesa",
    "Submissão de artigo",
    "Aguardando discente",
    "Aguardando orientador",
    "Finalizado"
]

PENDENCIAS_PADRAO = [
    "Definir tema",
    "Ajustar objetivos",
    "Revisar metodologia",
    "Folha de rosto",
    "Submissão ao CEP",
    "Coleta de dados",
    "Organizar banco de dados",
    "Análise estatística",
    "Escrita do projeto",
    "Escrita da dissertação/TCC",
    "Escrita do artigo",
    "Correção de referências",
    "Adequação às normas",
    "Agendar qualificação",
    "Agendar defesa",
    "Submeter artigo"
]

# =========================================================
# ESTILO VISUAL
# =========================================================

st.markdown(
    """
    <style>
    .main {background-color: #f7f8fb;}
    .block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
    div[data-testid="stMetricValue"] {font-size: 1.8rem;}
    .card {
        background: white;
        padding: 18px;
        border-radius: 18px;
        box-shadow: 0 3px 14px rgba(0,0,0,0.08);
        border: 1px solid #ececec;
    }
    .small-muted {color: #6b7280; font-size: 0.9rem;}
    </style>
    """,
    unsafe_allow_html=True
)

# =========================================================
# CONEXÃO COM GOOGLE SHEETS
# =========================================================

@st.cache_resource
def conectar_google_sheets():
    """
    Conecta ao Google Sheets usando as credenciais salvas no arquivo secrets.toml
    ou nos Secrets do Streamlit Cloud.
    """
    try:
        escopos = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]

        credenciais = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=escopos
        )

        cliente = gspread.authorize(credenciais)
        planilha = cliente.open(NOME_PLANILHA)
        return planilha

    except Exception as e:
        st.error("Não foi possível conectar ao Google Sheets.")
        st.warning(
            "Verifique se o arquivo secrets.toml foi configurado corretamente, "
            "se a planilha existe e se foi compartilhada com o e-mail da conta de serviço."
        )
        st.exception(e)
        st.stop()


def obter_aba(planilha, nome_aba, colunas):
    """Obtém uma aba existente ou cria a aba com cabeçalho padrão."""
    try:
        aba = planilha.worksheet(nome_aba)
    except gspread.WorksheetNotFound:
        aba = planilha.add_worksheet(title=nome_aba, rows=1000, cols=max(len(colunas), 10))
        aba.append_row(colunas)
    return aba


def ler_aba_com_cabecalho_padrao(aba, colunas):
    """
    Lê uma aba do Google Sheets usando o cabeçalho definido no código.
    Isso evita erro do gspread quando há cabeçalhos duplicados, vazios
    ou diferentes na primeira linha da planilha.
    """
    valores = aba.get_all_values()

    if not valores:
        aba.update([colunas])
        return pd.DataFrame(columns=colunas)

    primeira_linha = [str(x).strip() for x in valores[0]]

    if primeira_linha[:len(colunas)] != colunas:
        aba.update("A1", [colunas])
        valores = aba.get_all_values()

    dados = valores[1:]

    if not dados:
        return pd.DataFrame(columns=colunas)

    dados_ajustados = []
    for linha in dados:
        linha = linha[:len(colunas)]
        linha = linha + [""] * (len(colunas) - len(linha))
        dados_ajustados.append(linha)

    return pd.DataFrame(dados_ajustados, columns=colunas)


@st.cache_data(ttl=20)
def carregar_dados():
    """Carrega cadastro, registros e configurações da Planilha Google."""
    planilha = conectar_google_sheets()

    aba_cadastro = obter_aba(planilha, ABA_CADASTRO, COLUNAS_CADASTRO)
    aba_registros = obter_aba(planilha, ABA_REGISTROS, COLUNAS_REGISTROS)
    aba_config = obter_aba(planilha, ABA_CONFIG, ["Tipo", "Valor"])

    cadastro = ler_aba_com_cabecalho_padrao(aba_cadastro, COLUNAS_CADASTRO)
    registros = ler_aba_com_cabecalho_padrao(aba_registros, COLUNAS_REGISTROS)
    config = ler_aba_com_cabecalho_padrao(aba_config, ["Tipo", "Valor"])

    cadastro["Ativo"] = cadastro["Ativo"].replace("", "Sim")
    cadastro["Perfil"] = cadastro["Perfil"].replace("", "Orientando")

    if "Data" in registros.columns:
        registros["Data"] = pd.to_datetime(registros["Data"], errors="coerce").dt.date

    if "Prazo" in registros.columns:
        registros["Prazo"] = pd.to_datetime(registros["Prazo"], errors="coerce").dt.date

    situacoes = SITUACOES_PADRAO
    pendencias = PENDENCIAS_PADRAO

    if not config.empty and {"Tipo", "Valor"}.issubset(config.columns):
        sit = config.loc[config["Tipo"] == "Situação", "Valor"].dropna().astype(str).tolist()
        pen = config.loc[config["Tipo"] == "Pendência", "Valor"].dropna().astype(str).tolist()
        if sit:
            situacoes = sit
        if pen:
            pendencias = pen

    return cadastro, registros, situacoes, pendencias


def salvar_novo_registro(novo_registro):
    """Adiciona uma nova atualização na aba registros."""
    planilha = conectar_google_sheets()
    aba = obter_aba(planilha, ABA_REGISTROS, COLUNAS_REGISTROS)

    linha = [novo_registro.get(col, "") for col in COLUNAS_REGISTROS]
    aba.append_row(linha, value_input_option="USER_ENTERED")
    st.cache_data.clear()


def salvar_cadastro_completo(df_cadastro):
    """Substitui todo o cadastro na aba cadastro."""
    planilha = conectar_google_sheets()
    aba = obter_aba(planilha, ABA_CADASTRO, COLUNAS_CADASTRO)

    df = df_cadastro.copy()
    for col in COLUNAS_CADASTRO:
        if col not in df.columns:
            df[col] = ""

    df = df[COLUNAS_CADASTRO].fillna("")

    for col in df.columns:
        df[col] = df[col].astype(str)

    aba.clear()
    aba.update([COLUNAS_CADASTRO] + df.values.tolist())
    st.cache_data.clear()

# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================

def calcular_status_prazo(prazo):
    if pd.isna(prazo) or prazo in [None, ""]:
        return "Sem prazo"

    if isinstance(prazo, str):
        prazo = pd.to_datetime(prazo, errors="coerce")
        if pd.isna(prazo):
            return "Sem prazo"
        prazo = prazo.date()

    hoje = date.today()

    if prazo < hoje:
        return "Atrasado"

    dias = (prazo - hoje).days

    if dias <= 7:
        return "Próximo do prazo"

    return "Dentro do prazo"


def normalizar_email(email):
    if pd.isna(email):
        return ""
    return str(email).strip().lower()


def usuario_ativo(valor):
    return str(valor).strip().lower() in ["sim", "s", "yes", "1", "true"]


def resumo_ultima_situacao(df_registros, df_cadastro):
    if df_registros.empty:
        base = df_cadastro.copy()
        base["Última atualização"] = pd.NaT
        base["Situação"] = "Sem registro"
        base["Pendências"] = ""
        base["Responsável"] = ""
        base["Prazo"] = pd.NaT
        base["Prioridade"] = ""
        base["Dias sem atualização"] = None
        base["Status do prazo"] = "Sem prazo"
        return base

    df = df_registros.copy()
    df["Data_dt"] = pd.to_datetime(df["Data"], errors="coerce")

    ultimos = (
        df.sort_values("Data_dt")
        .groupby(["Discente", "Programa", "Nível"], as_index=False)
        .tail(1)
    )

    ultimos = ultimos.rename(columns={"Data": "Última atualização"})

    base = df_cadastro.merge(
        ultimos[
            [
                "Discente",
                "Programa",
                "Nível",
                "Última atualização",
                "Situação",
                "Pendências",
                "Prazo",
                "Prioridade",
                "Responsável"
            ]
        ],
        on=["Discente", "Programa", "Nível"],
        how="left"
    )

    hoje = pd.Timestamp(date.today())
    base["Última atualização_dt"] = pd.to_datetime(base["Última atualização"], errors="coerce")
    base["Dias sem atualização"] = (hoje - base["Última atualização_dt"]).dt.days

    base["Situação"] = base["Situação"].fillna("Sem registro")
    base["Pendências"] = base["Pendências"].fillna("")
    base["Prioridade"] = base["Prioridade"].fillna("")
    base["Responsável"] = base["Responsável"].fillna("")
    base["Status do prazo"] = base["Prazo"].apply(calcular_status_prazo)

    base = base.drop(columns=["Última atualização_dt"])
    return base

# =========================================================
# CARREGAMENTO
# =========================================================

df_cadastro, df_registros, SITUACOES, PENDENCIAS = carregar_dados()
df_resumo = resumo_ultima_situacao(df_registros, df_cadastro)

# =========================================================
# LOGIN INTERNO POR E-MAIL E SENHA
# =========================================================

st.sidebar.title("📚 Orientações")

email_usuario = st.sidebar.text_input(
    "Seu e-mail",
    placeholder="nome@email.com"
).strip().lower()

senha_usuario = st.sidebar.text_input(
    "Senha",
    type="password"
)

if not email_usuario or not senha_usuario:
    st.title("Controle de Orientações")
    st.info("Informe seu e-mail e senha na barra lateral para acessar o sistema.")
    st.stop()

if df_cadastro.empty:
    st.title("Controle de Orientações")
    st.error("Nenhum usuário cadastrado.")
    st.info("Cadastre pelo menos um orientador na aba cadastro da planilha Google.")
    st.stop()

if "Senha" not in df_cadastro.columns:
    st.title("Controle de Orientações")
    st.error("A coluna 'Senha' não foi encontrada na aba cadastro.")
    st.info("A primeira linha da aba cadastro deve ser: Discente | Programa | Nível | Email | Senha | Perfil | Ativo")
    st.stop()

if "Perfil" not in df_cadastro.columns:
    st.title("Controle de Orientações")
    st.error("A coluna 'Perfil' não foi encontrada na aba cadastro.")
    st.info("A primeira linha da aba cadastro deve ser: Discente | Programa | Nível | Email | Senha | Perfil | Ativo")
    st.stop()

df_cadastro["Email_normalizado"] = (
    df_cadastro["Email"]
    .astype(str)
    .str.strip()
    .str.lower()
)

usuario_logado = df_cadastro[
    (df_cadastro["Email_normalizado"] == email_usuario)
    &
    (df_cadastro["Senha"].astype(str).str.strip() == senha_usuario)
]

if usuario_logado.empty:
    st.title("Controle de Orientações")
    st.error("E-mail ou senha inválidos.")
    st.stop()

if not usuario_ativo(usuario_logado.iloc[0]["Ativo"]):
    st.title("Controle de Orientações")
    st.error("Usuário inativo.")
    st.stop()

nome_usuario = str(usuario_logado.iloc[0]["Discente"]).strip()
perfil_usuario = str(usuario_logado.iloc[0]["Perfil"]).strip()
acesso_orientador = perfil_usuario == "Orientador"

st.sidebar.success(f"Logado como: {nome_usuario}")
st.sidebar.caption(f"Perfil: {perfil_usuario}")

if not acesso_orientador:
    nomes_permitidos = usuario_logado[["Discente", "Programa", "Nível"]].drop_duplicates()

    df_resumo = df_resumo.merge(
        nomes_permitidos,
        on=["Discente", "Programa", "Nível"],
        how="inner"
    )

    df_registros = df_registros.merge(
        nomes_permitidos,
        on=["Discente", "Programa", "Nível"],
        how="inner"
    )

# =========================================================
# MENU E FILTROS
# =========================================================

if acesso_orientador:
    paginas = [
        "Painel geral",
        "Nova atualização",
        "Histórico",
        "Cadastro de orientandos",
        "Exportar dados"
    ]
else:
    paginas = [
        "Minha orientação",
        "Nova atualização",
        "Meu histórico"
    ]

pagina = st.sidebar.radio("Menu", paginas)

st.sidebar.divider()
st.sidebar.caption("Filtros")

programas = sorted(df_resumo["Programa"].dropna().unique())
niveis = sorted(df_resumo["Nível"].dropna().unique())

filtro_programa = st.sidebar.multiselect("Programa", programas, default=programas)
filtro_nivel = st.sidebar.multiselect("Nível", niveis, default=niveis)

resumo_filtrado = df_resumo[
    df_resumo["Programa"].isin(filtro_programa) &
    df_resumo["Nível"].isin(filtro_nivel)
].copy()

# =========================================================
# PAINEL GERAL DO ORIENTADOR
# =========================================================

if pagina == "Painel geral":
    st.title("Painel de Controle das Orientações")
    st.caption("Visão geral dos orientandos, situação atual, pendências e prazos.")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Orientandos", len(resumo_filtrado))
    col2.metric("Sem registro", int((resumo_filtrado["Situação"] == "Sem registro").sum()))
    col3.metric("Atrasados", int((resumo_filtrado["Status do prazo"] == "Atrasado").sum()))
    col4.metric("Sem atualização > 30 dias", int((resumo_filtrado["Dias sem atualização"].fillna(999) > 30).sum()))

    st.divider()

    col_a, col_b = st.columns([3, 1])

    with col_a:
        st.subheader("Situação atual dos orientandos")
        tabela = resumo_filtrado[
            [
                "Discente",
                "Programa",
                "Nível",
                "Email",
                "Situação",
                "Pendências",
                "Responsável",
                "Prazo",
                "Status do prazo",
                "Dias sem atualização",
                "Prioridade"
            ]
        ].sort_values(["Status do prazo", "Dias sem atualização"], ascending=[True, False])

        st.dataframe(tabela, use_container_width=True, hide_index=True, height=600)

    with col_b:
        st.subheader("Distribuição por nível")
        if not resumo_filtrado.empty:
            fig = px.pie(resumo_filtrado, names="Nível", hole=0.45)
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("Situação")
        contagem_situacao = resumo_filtrado["Situação"].value_counts().reset_index()
        contagem_situacao.columns = ["Situação", "Total"]
        fig2 = px.bar(contagem_situacao, x="Total", y="Situação", orientation="h")
        st.plotly_chart(fig2, use_container_width=True)

    st.divider()
    st.subheader("Alertas")

    atrasados = resumo_filtrado[resumo_filtrado["Status do prazo"] == "Atrasado"]
    parados = resumo_filtrado[resumo_filtrado["Dias sem atualização"].fillna(999) > 30]

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("### Prazos atrasados")
        if atrasados.empty:
            st.success("Nenhum prazo atrasado.")
        else:
            st.dataframe(
                atrasados[["Discente", "Situação", "Prazo", "Responsável"]],
                use_container_width=True,
                hide_index=True
            )

    with c2:
        st.markdown("### Sem atualização recente")
        if parados.empty:
            st.success("Nenhum orientando parado há mais de 30 dias.")
        else:
            st.dataframe(
                parados[["Discente", "Programa", "Nível", "Dias sem atualização"]],
                use_container_width=True,
                hide_index=True
            )

# =========================================================
# PAINEL DO ORIENTANDO
# =========================================================

elif pagina == "Minha orientação":
    st.title("Minha orientação")
    st.caption("Resumo da sua situação atual e das últimas atualizações.")

    if resumo_filtrado.empty:
        st.info("Não há dados cadastrados para este acesso.")
    else:
        registro = resumo_filtrado.iloc[0]

        col1, col2, col3 = st.columns(3)
        col1.metric("Situação", registro["Situação"])
        col2.metric("Status do prazo", registro["Status do prazo"])
        dias = registro["Dias sem atualização"]
        col3.metric("Dias sem atualização", "Sem registro" if pd.isna(dias) else int(dias))

        st.divider()
        st.subheader("Resumo atual")
        st.write(f"**Discente:** {registro['Discente']}")
        st.write(f"**Programa:** {registro['Programa']}")
        st.write(f"**Nível:** {registro['Nível']}")
        st.write(f"**Pendências:** {registro['Pendências']}")
        st.write(f"**Responsável pela próxima ação:** {registro['Responsável']}")
        st.write(f"**Prazo:** {registro['Prazo']}")
        st.write(f"**Prioridade:** {registro['Prioridade']}")

# =========================================================
# NOVA ATUALIZAÇÃO
# =========================================================

elif pagina == "Nova atualização":
    st.title("Nova atualização de orientação")
    st.caption("Registre situação atual, pendências e observações.")

    if acesso_orientador:
        orientandos_ativos = df_cadastro[
            df_cadastro["Ativo"].apply(usuario_ativo)
            &
            (df_cadastro["Perfil"].astype(str).str.strip() == "Orientando")
        ].copy()
    else:
        orientandos_ativos = usuario_logado.copy()

    orientandos_ativos["Nome completo"] = (
        orientandos_ativos["Discente"].astype(str)
        + " — "
        + orientandos_ativos["Programa"].astype(str)
        + " — "
        + orientandos_ativos["Nível"].astype(str)
    )

    if orientandos_ativos.empty:
        st.warning("Nenhum orientando ativo encontrado.")
        st.stop()

    with st.form("form_nova_atualizacao", clear_on_submit=True):
        col1, col2 = st.columns(2)

        with col1:
            escolha = st.selectbox("Discente", orientandos_ativos["Nome completo"].tolist())
            data_registro = st.date_input("Data da atualização", value=date.today())
            situacao = st.selectbox("Situação atual", SITUACOES)
            prioridade = st.selectbox("Prioridade", ["Baixa", "Média", "Alta"])

        with col2:
            pendencias_escolhidas = st.multiselect("Pendências", PENDENCIAS)
            responsavel = st.selectbox("Responsável pela próxima ação", ["Discente", "Orientador", "Ambos", "Outro"])
            prazo = st.date_input("Prazo da próxima entrega", value=None)

        observacoes = st.text_area(
            "Observações",
            height=140,
            placeholder="Ex.: enviei nova versão do projeto; aguardando parecer; combinar análise dos dados..."
        )

        enviar = st.form_submit_button("Salvar atualização", type="primary")

    if enviar:
        linha = orientandos_ativos[orientandos_ativos["Nome completo"] == escolha].iloc[0]

        novo = {
            "Data": data_registro.strftime("%Y-%m-%d"),
            "Hora": datetime.now().strftime("%H:%M:%S"),
            "Discente": linha["Discente"],
            "Programa": linha["Programa"],
            "Nível": linha["Nível"],
            "Email": linha.get("Email", ""),
            "Situação": situacao,
            "Pendências": "; ".join(pendencias_escolhidas),
            "Responsável": responsavel,
            "Prazo": prazo.strftime("%Y-%m-%d") if prazo else "",
            "Prioridade": prioridade,
            "Observações": observacoes,
            "Atualizado_por": email_usuario
        }

        salvar_novo_registro(novo)
        st.success("Atualização salva com sucesso.")

# =========================================================
# HISTÓRICO
# =========================================================

elif pagina in ["Histórico", "Meu histórico"]:
    st.title("Histórico de orientações")

    if df_registros.empty:
        st.info("Ainda não há registros de orientação.")
    else:
        df_hist = df_registros.copy()
        df_hist = df_hist[
            df_hist["Programa"].isin(filtro_programa) &
            df_hist["Nível"].isin(filtro_nivel)
        ]

        if acesso_orientador:
            nomes = sorted(df_hist["Discente"].dropna().unique())
            nome_filtro = st.multiselect("Filtrar por discente", nomes, default=nomes)
            situacao_filtro = st.multiselect("Filtrar por situação", sorted(df_hist["Situação"].dropna().unique()))

            df_hist = df_hist[df_hist["Discente"].isin(nome_filtro)]
            if situacao_filtro:
                df_hist = df_hist[df_hist["Situação"].isin(situacao_filtro)]

        df_hist = df_hist.sort_values(["Data", "Hora"], ascending=False)
        st.dataframe(df_hist, use_container_width=True, hide_index=True, height=600)

        st.subheader("Linha do tempo")

        if acesso_orientador:
            nomes_timeline = sorted(df_hist["Discente"].dropna().unique())
            if nomes_timeline:
                discente_timeline = st.selectbox("Selecionar discente", nomes_timeline)
                linha_tempo = df_hist[df_hist["Discente"] == discente_timeline].sort_values(["Data", "Hora"], ascending=False)
            else:
                linha_tempo = pd.DataFrame(columns=df_hist.columns)
        else:
            linha_tempo = df_hist.sort_values(["Data", "Hora"], ascending=False)

        for _, row in linha_tempo.iterrows():
            with st.expander(f"{row['Data']} — {row['Situação']}"):
                st.write(f"**Hora:** {row['Hora']}")
                st.write(f"**Discente:** {row['Discente']}")
                st.write(f"**Programa:** {row['Programa']}")
                st.write(f"**Nível:** {row['Nível']}")
                st.write(f"**Pendências:** {row['Pendências']}")
                st.write(f"**Responsável:** {row['Responsável']}")
                st.write(f"**Prazo:** {row['Prazo']}")
                st.write(f"**Prioridade:** {row['Prioridade']}")
                st.write(f"**Observações:** {row['Observações']}")
                st.write(f"**Atualizado por:** {row['Atualizado_por']}")

# =========================================================
# CADASTRO DE ORIENTANDOS
# =========================================================

elif pagina == "Cadastro de orientandos":
    st.title("Cadastro de orientandos")
    st.caption("Inclua e edite os usuários autorizados a acessar o sistema.")

    st.info(
        "Para que o usuário acesse o sistema, cadastre e-mail, senha, perfil e status ativo. "
        "Use Perfil = Orientador para acesso completo e Perfil = Orientando para acesso individual."
    )

    with st.form("form_cadastro"):
        col1, col2, col3 = st.columns(3)

        with col1:
            novo_nome = st.text_input("Nome / Discente")
            novo_email = st.text_input("E-mail")
            nova_senha = st.text_input("Senha", type="password")

        with col2:
            novo_programa = st.selectbox("Programa", ["PPGSC", "PPGASFAR", "PRMU", "Farmácia", "Orientador", "Outro"])
            novo_nivel = st.selectbox("Nível", ["Doutorado", "Mestrado", "Residência", "Graduação", "Orientador", "Outro"])

        with col3:
            perfil = st.selectbox("Perfil", PERFIS)
            ativo = st.selectbox("Ativo", ["Sim", "Não"])

        cadastrar = st.form_submit_button("Adicionar usuário", type="primary")

    if cadastrar:
        if not novo_nome.strip():
            st.warning("Informe o nome.")
        elif not novo_email.strip():
            st.warning("Informe o e-mail.")
        elif not nova_senha.strip():
            st.warning("Informe uma senha.")
        else:
            novo_cad = pd.DataFrame([
                {
                    "Discente": novo_nome.strip(),
                    "Programa": novo_programa,
                    "Nível": novo_nivel,
                    "Email": novo_email.strip().lower(),
                    "Senha": nova_senha.strip(),
                    "Perfil": perfil,
                    "Ativo": ativo
                }
            ])

            df_cadastro = pd.concat([df_cadastro[COLUNAS_CADASTRO], novo_cad], ignore_index=True)
            salvar_cadastro_completo(df_cadastro)
            st.success("Usuário cadastrado com sucesso.")

    st.subheader("Base atual")
    editado = st.data_editor(
        df_cadastro[COLUNAS_CADASTRO],
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic"
    )

    if st.button("Salvar alterações no cadastro"):
        salvar_cadastro_completo(editado)
        st.success("Cadastro atualizado.")

# =========================================================
# EXPORTAÇÃO
# =========================================================

elif pagina == "Exportar dados":
    st.title("Exportar dados")

    resumo_export = df_resumo.to_csv(index=False).encode("utf-8-sig")
    registros_export = df_registros.to_csv(index=False).encode("utf-8-sig")
    cadastro_export = df_cadastro[COLUNAS_CADASTRO].to_csv(index=False).encode("utf-8-sig")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button(
            "Baixar resumo atual",
            data=resumo_export,
            file_name="resumo_orientacoes.csv",
            mime="text/csv"
        )

    with col2:
        st.download_button(
            "Baixar histórico completo",
            data=registros_export,
            file_name="historico_orientacoes.csv",
            mime="text/csv"
        )

    with col3:
        st.download_button(
            "Baixar cadastro",
            data=cadastro_export,
            file_name="cadastro_orientandos.csv",
            mime="text/csv"
        )

    st.info("Os dados principais ficam salvos na Planilha Google vinculada ao aplicativo.")

