import streamlit as st
import pandas as pd
import plotly.express as px
import os
import unicodedata

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Dashboard Migrações Fortigate", layout="wide")

# --- FUNÇÕES UTILITÁRIAS ---
def _normalize_text(s):
    if pd.isna(s) or s == "":
        return ""
    s = str(s).lower()
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return s.strip()

# --- CARREGAMENTO DE DADOS ---
@st.cache_data
def load_data():
    arquivo_excel = 'ControleDeRevisitas.xlsx'
    arquivo_csv = 'ControleDeRevisitas.xlsx - CONTROLE.csv'

    df = None
    
    # 1. Tenta carregar Excel (Prioridade)
    if os.path.exists(arquivo_excel):
        try:
            df = pd.read_excel(arquivo_excel, engine='openpyxl')
        except Exception as e:
            st.warning(f"Aviso: Erro ao ler Excel direto ({e}). Tentando CSV...")

    # 2. Tenta carregar CSV se Excel falhou
    if df is None and os.path.exists(arquivo_csv):
        try:
            df = pd.read_csv(arquivo_csv, encoding='utf-8')
        except:
            try:
                df = pd.read_csv(arquivo_csv, encoding='latin1', sep=';')
            except:
                df = pd.read_csv(arquivo_csv, encoding='latin1')
            
    if df is None:
        return None

    # Limpeza básica de colunas
    df.columns = df.columns.str.strip().str.replace('  ', ' ')

    # Normalizar datas
    col_1a_visita = '1º Visita'
    if col_1a_visita not in df.columns:
        cols = [c for c in df.columns if 'Visita' in c and '1' in c]
        if cols: col_1a_visita = cols[0]

    date_cols = [col_1a_visita, '2º Visita', '3º Visita']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    return df, col_1a_visita

# --- CSS PARA FIXAR O CABEÇALHO ---
st.markdown("""
    <style>
    /* A mágica acontece aqui:
       Procuramos um bloco vertical que contenha a nossa DIV marcadora com id="main-header-marker".
       Assim, o CSS só afeta este bloco específico e ignora as métricas lá de baixo.
    */
    div[data-testid="stVerticalBlock"] > div:has(div#main-header-marker) {
        position: sticky;
        top: 0;
        background-color: #0E1117; /* Garante fundo opaco */
        z-index: 1;
        padding-top: 14px; /* Espaço extra para o texto não cortar */
        padding-bottom: 11px; /* Espaço extra para o texto não cortar */
    }
    div:has(div#b3) {
        margin-top: 0px; /* Espaço acima do conteúdo rolável */
        z-index: 1015;
    }
    </style>
""", unsafe_allow_html=True)

# --- MOTOR DE CLASSIFICAÇÃO (POR VISITA) ---
def classificar_evento_isolado(motivo, obs, data_visita, visita_anterior_concluida=False):
    """
    Classifica o que aconteceu em uma visita específica.
    """
    if visita_anterior_concluida:
        return "N/A (Já Concluído)"

    m = _normalize_text(str(motivo))
    o = _normalize_text(str(obs))
    
    # 1. Cancelado (Prioridade Global na linha)
    if 'cancelad' in m or 'cancelad' in o:
        if not ('reagend' in m or 'reagend' in o or 'remarca' in m or 'remarca' in o):
            return "Cancelado"

    # 2. Não Realizada
    if pd.isna(data_visita):
        return "Não Realizada"

    # 3. Concluído
    if 'conclui' in m or 'finaliz' in m or 'migrada' in m:
        return "Concluído"

    # 4. Falhas Específicas (Pendências)
    if 'misto' in m or ('mvc' in m and 'telebras' in m):
        return "Misto (Telebras + MVC)"
    
    if 'telebras' in m or 'infra' in m or 'link' in m or 'tlb' in m:
        return "Infraestrutura Telebras"
    
    if 'mvc' in m or 'operacional' in m or 'fotos' in m or 'doc' in m:
        return "Pendência Operacional (MVC)"
    
    if 'acesso' in m or 'ma' in m or 'logistica' in m or 'agend' in m:
        return "Acesso / MA / Logística"

    # 5. Indefinido (Mas tem data) - NÃO É PENDÊNCIA OPERACIONAL
    if len(m) < 3 and len(o) > 3:
        return "A Verificar (Ler Obs)"
    
    return "A Verificar (Bitrix/Teams)"

# --- LÓGICA DE STATUS FINAL ---
def calcular_status_final(row):
    s1 = row['Status_V1']
    s2 = row['Status_V2']
    s3 = row['Status_V3']

    # 1. Checagem de Cancelamento (Prioridade Total)
    if "Cancelado" in [s1, s2, s3]:
        return "Cancelado"

    # 2. Definição do status baseado na ÚLTIMA visita válida realizada
    status_atual = "Não Iniciado"
    
    # Verifica de trás para frente (V3 -> V2 -> V1)
    if s3 not in ["Não Realizada", "N/A (Já Concluído)"]:
        status_atual = s3
    elif s2 not in ["Não Realizada", "N/A (Já Concluído)"]:
        status_atual = s2
    elif s1 not in ["Não Realizada", "N/A (Já Concluído)"]:
        status_atual = s1
    
    return status_atual

# --- INTERFACE GRÁFICA ---
data_load = load_data()

if data_load is None:
    st.error("❌ Arquivo de dados não encontrado.")
else:
    df, col_1a = data_load

    # --- PROCESSAMENTO ---
    # Classificar cada visita individualmente
    df['Status_V1'] = df.apply(lambda row: classificar_evento_isolado(
        row.get('Motivo_Padronizado'), row.get('Obs'), row[col_1a]
    ), axis=1)

    df['Status_V2'] = df.apply(lambda row: classificar_evento_isolado(
        row.get('Motivo_Padronizado2'), row.get('Obs2'), row.get('2º Visita'), 
        visita_anterior_concluida=(row['Status_V1'] == 'Concluído')
    ), axis=1)

    df['Status_V3'] = df.apply(lambda row: classificar_evento_isolado(
        row.get('Motivo_Padronizado3'), row.get('Obs3'), row.get('3º Visita'), 
        visita_anterior_concluida=(row['Status_V2'] == 'Concluído' or row['Status_V1'] == 'Concluído')
    ), axis=1)

    # Calcular Status Atual
    df['Status_Final'] = df.apply(calcular_status_final, axis=1)
    df['Mes_Inicial'] = df[col_1a].dt.to_period('M').astype(str)

    # --- SIDEBAR E FILTROS ---
    st.sidebar.title("Filtros")
    # Filtra apenas meses válidos (exclui NaT/Não Iniciado da lista de seleção)
    meses_validos = sorted(df[df['Mes_Inicial'] != 'NaT']['Mes_Inicial'].unique().astype(str))
    mes_sel = st.sidebar.multiselect("Mês da 1ª Visita", meses_validos, default=meses_validos)
    
    # Se a seleção estiver vazia (usuário removeu tudo), o DF ficará vazio
    df_filtrado = df[df['Mes_Inicial'].isin(mes_sel)].copy()

    

    # Mapa de Cores Padronizado
    color_map = {
        'Concluído': '#2ecc71', 
        'Infraestrutura Telebras': '#e74c3c', 
        'Pendência Operacional (MVC)': '#e67e22', 
        'Misto (Telebras + MVC)': '#d35400',
        'Acesso / MA / Logística': '#f1c40f', 
        'A Verificar (Bitrix/Teams)': '#95a5a6', 
        'A Verificar (Ler Obs)': '#95a5a6',
        'Cancelado': '#34495e', 
        'Não Realizada': '#ecf0f1', 
        'Não Iniciado': '#bdc3c7'
    }

    # --- DASHBOARD ---
    
    st.title("📊 Dashboard Analítico: Ativações SD-WAN")

    # 1. CÁLCULO DE KPIS (REGRA ESTRITA)
    total = len(df_filtrado)
    finalizados = len(df_filtrado[df_filtrado['Status_Final'] == 'Concluído'])
    cancelados = len(df_filtrado[df_filtrado['Status_Final'] == 'Cancelado'])

    # Definição estrita do que é Pendência
    lista_pendencias = [
        'Infraestrutura Telebras', 
        'Pendência Operacional (MVC)', 
        'Misto (Telebras + MVC)', 
        'Acesso / MA / Logística'
    ]

    pendentes_reais = len(df_filtrado[df_filtrado['Status_Final'].isin(lista_pendencias)])

    outros = total - finalizados - cancelados - pendentes_reais

    # --- 1. CABEÇALHO FIXO (Sticky) ---
    with st.container():
        # Marca invisível que o CSS procura para aplicar o sticky somente a este container
        st.markdown('<div id="main-header-marker"></div>'
                    '<div id="b3"></div>'
                    , unsafe_allow_html=True)

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total no Período", total)

        # CORREÇÃO DO ERRO DE DIVISÃO POR ZERO
        pct_concluidas = (finalizados/total*100) if total > 0 else 0.0
        k2.metric("Concluídas", finalizados, delta=f"{pct_concluidas:.1f}%")

        # Mostrar apenas o número no header (o botão fica no conteúdo rolável abaixo)
        k3.metric("Pendências Atuais", pendentes_reais, delta_color="inverse", help="Infra, MVC ou Acesso.")
        k4.metric("Cancelados", cancelados)

    # --- 2. CONTEÚDO QUE ROLA (botão e demais elementos) ---
    b1, b2, b3, b4 = st.columns(4)
    with b3:
        if pendentes_reais > 0:
            with st.popover("🔍 Ver Detalhes", use_container_width=False):
                st.subheader("Detalhamento das Pendências Atuais")
                df_detalhe = df_filtrado[df_filtrado['Status_Final'].isin(lista_pendencias)]
                counts_detalhe = df_detalhe['Status_Final'].value_counts().reset_index()
                counts_detalhe.columns = ['Tipo de Pendência', 'Quantidade']
                st.dataframe(counts_detalhe, use_container_width=True, hide_index=True)

                # Mini gráfico no popover
                fig_mini = px.pie(counts_detalhe, names='Tipo de Pendência', values='Quantidade', hole=0.5, color='Tipo de Pendência', color_discrete_map=color_map)
                fig_mini.update_layout(height=250, margin=dict(t=0, b=0, l=0, r=0))
                st.plotly_chart(fig_mini, use_container_width=True)

    # --- Visão Geral (Gráfico de Pizza) ---
    # Reaproveitar df_filtrado para distribuir por Status_Final
    pie_counts = df_filtrado['Status_Final'].value_counts().reset_index()
    pie_counts.columns = ['Status', 'Quantidade'] 

    pie_color_map = {
        'Concluído': '#2ecc71', 
        'Infraestrutura Telebras': '#e74c3c', 
        'Pendência Operacional (MVC)': '#e67e22', 
        'Misto (Telebras + MVC)': '#d35400',
        'Acesso / MA / Logística': '#f1c40f', 
        'A Verificar (Bitrix/Teams)': '#95a5a6', 
        'A Verificar (Ler Obs)': '#95a5a6',
        'Cancelado': '#34495e', 
        'Não Realizada': '#ecf0f1', 
        'Não Iniciado': '#bdc3c7'
    }

    fig_overall_pie = px.pie(
        pie_counts,
        names='Status',
        values='Quantidade',
        title='Visão Geral: Distribuição por Status',
        hole=0.35,
        color_discrete_map=pie_color_map
    )
    st.plotly_chart(fig_overall_pie, use_container_width=True)

    if total == 0:
        st.warning("⚠️ Nenhum dado encontrado para os filtros selecionados. Selecione pelo menos um mês na barra lateral.")
        st.stop() # Interrompe a execução aqui para não quebrar os gráficos abaixo

    st.divider()

    # 2. ANÁLISE DA 1ª VISITA
    st.header("1️⃣ Análise da 1ª Visita")
    v1_stats = df_filtrado['Status_V1'].value_counts().reset_index()
    v1_stats.columns = ['Status', 'Quantidade']
    
    # Métricas V1
    c1, c2, c3, c4 = st.columns(4)
    # Proteção com .get para evitar erro se não houver o status
    def get_count(df_stats, status_name):
        res = df_stats[df_stats['Status'] == status_name]['Quantidade'].sum()
        return int(res)

    c1.metric("Concluídas na 1ª", get_count(v1_stats, 'Concluído'))
    c2.metric("Falha Infra Telebras", get_count(v1_stats, 'Infraestrutura Telebras'), delta_color="inverse")
    c3.metric("Falha MVC", get_count(v1_stats, 'Pendência Operacional (MVC)'), delta_color="inverse")
    c4.metric("Canceladas", get_count(v1_stats, 'Cancelado'))
    
    fig_v1 = px.bar(
        v1_stats, x='Quantidade', y='Status', orientation='h', 
        text_auto=True, color='Status', color_discrete_map=color_map,
        title="Resultados da 1ª Visita"
    )
    st.plotly_chart(fig_v1, use_container_width=True)

    st.divider()

    # 3. RESOLVIDO NA 2ª VISITA
    st.header("2️⃣ RESOLVIDO NA 2ª VISITA")

    # Filtro: Quem foi resolvido na V2
    concluidos_v2 = df_filtrado[df_filtrado['Status_V2'] == 'Concluído']
    qtd_v2_ok = len(concluidos_v2)
    
    st.subheader(f"✅ {qtd_v2_ok} localidades concluídas na 2ª tentativa")
    
    if qtd_v2_ok > 0:
        # Mostra a causa original (V1)
        origem_v2 = concluidos_v2['Status_V1'].value_counts().reset_index()
        origem_v2.columns = ['Motivo da Falha Original (V1)', 'Qtd Resolvida']
        
        col_v2_g, col_v2_t = st.columns([2, 1])
        with col_v2_g:
            fig_v2 = px.bar(
                origem_v2, x='Qtd Resolvida', y='Motivo da Falha Original (V1)', orientation='h',
                text_auto=True, color='Motivo da Falha Original (V1)', color_discrete_map=color_map,
                title="Causa Raiz das localidades recuperadas na 2ª Visita"
            )
            st.plotly_chart(fig_v2, use_container_width=True)
        with col_v2_t:
            st.markdown("**Detalhamento:**")
            st.dataframe(origem_v2, use_container_width=True, hide_index=True)

    st.divider()

    # 4. RESOLVIDO NA 3ª VISITA (Visual idêntico à seção 2)
    st.header("3️⃣ RESOLVIDO NA 3ª VISITA")
    
    concluidos_v3 = df_filtrado[df_filtrado['Status_V3'] == 'Concluído']
    qtd_v3_ok = len(concluidos_v3)
    
    st.subheader(f"✅ {qtd_v3_ok} localidades concluídas na 3ª tentativa")
    
    if qtd_v3_ok > 0:
        # Mantendo o padrão: Mostra a Causa Raiz (V1) para entender a origem do problema persistente
        origem_v3 = concluidos_v3['Status_V1'].value_counts().reset_index()
        origem_v3.columns = ['Motivo da Falha Original (V1)', 'Qtd Resolvida']
        
        col_v3_g, col_v3_t = st.columns([2, 1])
        with col_v3_g:
            fig_v3 = px.bar(
                origem_v3, x='Qtd Resolvida', y='Motivo da Falha Original (V1)', orientation='h',
                text_auto=True, color='Motivo da Falha Original (V1)', color_discrete_map=color_map,
                title="Causa Raiz das localidades recuperadas na 3ª Visita"
            )
            st.plotly_chart(fig_v3, use_container_width=True)
        with col_v3_t:
            st.markdown("**Detalhamento:**")
            st.dataframe(origem_v3, use_container_width=True, hide_index=True)

    st.divider()

    # --- 5. EXPORTAÇÃO DE DADOS ---
    st.subheader("📋 Histórico Completo & Priorização")
    st.markdown("A tabela abaixo apresenta **apenas as pendências**, ordenadas por prioridade: **1. MVC/Acesso (Laranja)** -> **2. Telebras (Azul)** -> **Mais Antigas**.")
    
    colunas_export = [
        'SITE-ID', 'LOCALIDADE', 'Mes_Inicial',
        'Status_V1', 'Status_V2', 'Status_V3', 
        'Status_Final', 'Obs', 'Obs2', 'Obs3'
    ]
    cols_to_export = [c for c in colunas_export if c in df_filtrado.columns]
    
    # 1. Definir grupos de prioridade
    grupo_laranja = ['Pendência Operacional (MVC)', 'Acesso / MA / Logística', 'Misto (Telebras + MVC)']
    grupo_azul = ['Infraestrutura Telebras']
    
    # 2. Criar função de ordenação
    def get_prioridade(status):
        if status in grupo_laranja:
            return 0 # Prioridade Máxima
        if status in grupo_azul:
            return 1 # Prioridade Secundária
        return 2 # Outros
    
    df_filtrado['Prioridade_Sort'] = df_filtrado['Status_Final'].apply(get_prioridade)
    
    # 3. Filtrar para mostrar APENAS as pendências na visualização (conforme pedido)
    # Lista de todos os itens considerados pendência
    todos_pendentes = grupo_laranja + grupo_azul
    df_pendencias = df_filtrado[df_filtrado['Status_Final'].isin(todos_pendentes)].copy()

    # 4. Ordenar: Prioridade (0, 1) -> Data Antiga para Nova
    # col_1a é o nome da coluna de data (ex: '1º Visita')
    df_sorted = df_pendencias.sort_values(by=['Prioridade_Sort', col_1a], ascending=[True, True])
    
    # 5. Prepara DF final
    df_visual = df_sorted[cols_to_export].copy()

    # 6. Função de Estilo (Highlight) - Cores Solicitadas
    def highlight_priorities(row):
        status = row['Status_Final']
        
        # Laranja para MVC / Acesso / Logística
        if status in grupo_laranja:
            return ['background-color: #ffccbc; color: black'] * len(row) # Laranja claro
        
        # Azul para Telebras
        if status in grupo_azul:
            return ['background-color: #bbdefb; color: black'] * len(row) # Azul claro
            
        return [''] * len(row)

    # Exibe a tabela estilizada
    st.dataframe(
        df_visual.style.apply(highlight_priorities, axis=1), 
        use_container_width=True, 
        hide_index=True
    )

    if not df_visual.empty:
        csv = df_visual.to_csv(index=False).encode('utf-8-sig')
        
        st.download_button(
            data=csv,
            file_name="Lista_Pendencias_Priorizada.csv",
            mime="text/csv",
            label="📥 Baixar Lista de Pendências Priorizada (CSV)",
        )