import streamlit as st
import pandas as pd
import datetime
import plotly.express as px
import sys
import os

# =================================================================
# 1. CONFIGURAÇÃO DA PÁGINA (Deve ser a primeira linha)
# =================================================================
st.set_page_config(
    page_title="Dashboard Delivery",
    page_icon="🏍️",
    layout="wide"
)

# =================================================================
# 2. FUNÇÕES DE CARREGAMENTO E CÁLCULO
# =================================================================

@st.cache_data
def carregar_dados():
    # Lista de tentativas: onde o arquivo pode estar?
    caminhos_possiveis = [
        # 1. TENTATIVA IDEAL: No caminho fixo do Google Drive (o arquivo VIVO)
        r"G:\Meu Drive\Delivery\Delivery\delivery_atualizado.xlsx",
        
        # 2. TENTATIVA ALTERNATIVA: Vai que o Drive da pessoa é H: ou I:?
        r"H:\Meu Drive\Delivery\Delivery\delivery_atualizado.xlsx",
        
        # 3. TENTATIVA LOCAL: Se a pessoa colocou o .exe NA MESMA PASTA do Excel
        os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)), "delivery_atualizado.xlsx")
    ]

    caminho_final = None
    
    # O loop testa caminho por caminho até achar um que existe
    for caminho in caminhos_possiveis:
        if os.path.exists(caminho):
            caminho_final = caminho
            break
            
    try:
        if caminho_final is None:
            # Se não achou em lugar nenhum, mostra erro e pede ajuda
            st.error("❌ ERRO CRÍTICO: Não encontrei a base de dados 'delivery_atualizado.xlsx'!")
            st.warning("Certifique-se de que você tem o Google Drive instalado (G:) ou coloque este executável na mesma pasta do arquivo Excel.")
            return pd.DataFrame()

        # Carrega o arquivo encontrado
        df = pd.read_excel(caminho_final, sheet_name='Dados Gerais')
        
        # Limpezas
        df.columns = df.columns.str.strip()
        if 'LOJAS' in df.columns:
            df['LOJAS'] = df['LOJAS'].str.strip().str.upper()
        df['Data de entrega'] = pd.to_datetime(df['Data de entrega'], errors='coerce')
        df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce')
        
        return df
        
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        return pd.DataFrame()

def calcular_metricas(df):
    hoje = pd.Timestamp.now().normalize()
    
    # --- SEMANAL ---
    idx_domingo = (hoje.weekday() + 1) % 7
    ultimo_domingo = hoje - pd.Timedelta(days=idx_domingo)
    fim_sem = ultimo_domingo
    inicio_sem = fim_sem - pd.Timedelta(days=6)
    
    inicio_sem_ly = inicio_sem - pd.Timedelta(days=364)
    fim_sem_ly = fim_sem - pd.Timedelta(days=364)
    
    fat_semanal = df.loc[(df['Data de entrega'] >= inicio_sem) & (df['Data de entrega'] <= fim_sem), 'Valor'].sum()
    fat_semanal_ly = df.loc[(df['Data de entrega'] >= inicio_sem_ly) & (df['Data de entrega'] <= fim_sem_ly), 'Valor'].sum()
    
    # --- MENSAL ---
    if hoje.day <= 5:
        primeiro_dia_mes = (hoje.replace(day=1) - pd.Timedelta(days=1)).replace(day=1)
        fim_mes = hoje.replace(day=1) - pd.Timedelta(days=1)
    else:
        primeiro_dia_mes = hoje.replace(day=1)
        fim_mes = hoje

    inicio_mes_ly = primeiro_dia_mes - pd.DateOffset(years=1)
    fim_mes_ly = fim_mes - pd.DateOffset(years=1)

    fat_mensal = df.loc[(df['Data de entrega'] >= primeiro_dia_mes) & (df['Data de entrega'] <= fim_mes), 'Valor'].sum()
    fat_mensal_ly = df.loc[(df['Data de entrega'] >= inicio_mes_ly) & (df['Data de entrega'] <= fim_mes_ly), 'Valor'].sum()

    # --- ANUAL ---
    inicio_ano = hoje.replace(month=1, day=1)
    inicio_ano_ly = inicio_ano - pd.DateOffset(years=1)
    fim_ano_ly = hoje - pd.DateOffset(years=1)
    
    fat_anual = df.loc[(df['Data de entrega'] >= inicio_ano) & (df['Data de entrega'] <= hoje), 'Valor'].sum()
    fat_anual_ly = df.loc[(df['Data de entrega'] >= inicio_ano_ly) & (df['Data de entrega'] <= fim_ano_ly), 'Valor'].sum()
    
    return (fat_semanal, fat_semanal_ly, fat_mensal, fat_mensal_ly, fat_anual, fat_anual_ly, inicio_sem, fim_sem)

def plotar_grafico_comparativo(valor_atual, valor_ly):
    """Gera gráfico de barras finas e coladas"""
    dados_grafico = pd.DataFrame({
        'Periodo': ['Ano Anterior (LY)', 'Atual'],
        'Faturamento': [valor_ly, valor_atual],
        'Grupo': ['Comparativo', 'Comparativo']
    })
    
    fig = px.bar(
        dados_grafico, x='Grupo', y='Faturamento', text_auto='.2s', 
        color='Periodo', barmode='group',
        color_discrete_map={'Ano Anterior (LY)': '#d3d3d3', 'Atual': '#0068c9'}
    )
    
    fig.update_layout(
        showlegend=False, margin=dict(t=10, b=0, l=0, r=0), height=250, 
        plot_bgcolor="rgba(0,0,0,0)",
        yaxis={'showgrid': False, 'visible': False}, 
        xaxis={'showgrid': False, 'visible': False},
        bargap=0.7, bargroupgap=0
    )
    fig.update_traces(textposition='outside', cliponaxis=False)
    return fig

# =================================================================
# 3. CONTROLE DE ESTADO (SESSION STATE)
# =================================================================
# Isso garante que a página "lembre" qual botão foi clicado
if 'visao_atual' not in st.session_state:
    st.session_state['visao_atual'] = 'indicadores' # Visão padrão ao abrir

# =================================================================
# 4. BARRA LATERAL (MENU)
# =================================================================
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/2979/2979698.png", width=100)
st.sidebar.title("Menu Gerencial")

# Botão 1: Muda o estado para 'indicadores'
if st.sidebar.button("📊 Indicadores Gerais", use_container_width=True):
    st.session_state['visao_atual'] = 'indicadores'

# Botão 2: Muda o estado para 'rankings'
if st.sidebar.button("🏆 Rankings e Lojas", use_container_width=True):
    st.session_state['visao_atual'] = 'rankings'

st.sidebar.markdown("---")
st.sidebar.info(f"Visualizando: **{st.session_state['visao_atual'].upper()}**")


# =================================================================
# 5. ÁREA PRINCIPAL
# =================================================================
st.title("🚀 Dashboard de Performance - Delivery")
st.markdown("---")

df = carregar_dados()

if not df.empty:
    
    # 1. Calcula as datas e métricas básicas
    (fat_semanal, fat_semanal_ly, 
     fat_mensal, fat_mensal_ly, 
     fat_anual, fat_anual_ly, 
     inicio_sem, fim_sem) = calcular_metricas(df)

    # 2. CÁLCULO DE DATAS MENSAIS (Adicione este bloco aqui para corrigir o erro)
    # Precisamos disso aqui fora para usar no "help" dos gráficos
    hoje = pd.Timestamp.now().normalize()
    if hoje.day <= 5:
        primeiro_dia_mes = (hoje.replace(day=1) - pd.Timedelta(days=1)).replace(day=1)
        fim_mes = hoje.replace(day=1) - pd.Timedelta(days=1)
    else:
        primeiro_dia_mes = hoje.replace(day=1)
        fim_mes = hoje

    # -----------------------------------------------------------
    # VISÃO 1: INDICADORES (KPIs)
    # -----------------------------------------------------------
    if st.session_state['visao_atual'] == 'indicadores':
        
        st.subheader("Resumo Financeiro vs. Ano Anterior (LY)")
        
        def calcular_var(atual, anterior):
            return ((atual - anterior) / anterior * 100) if anterior > 0 else 0.0

        var_sem = calcular_var(fat_semanal, fat_semanal_ly)
        var_mes = calcular_var(fat_mensal, fat_mensal_ly)
        var_ano = calcular_var(fat_anual, fat_anual_ly)

        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric(
                label="Faturamento Semanal", 
                value=f"R$ {fat_semanal:,.2f}", 
                delta=f"{var_sem:+.1f}% (LY: {fat_semanal_ly:,.0f})",
                help=f"Período: {inicio_sem.strftime('%d/%m')} até {fim_sem.strftime('%d/%m')}"
            )
            st.plotly_chart(plotar_grafico_comparativo(fat_semanal, fat_semanal_ly), use_container_width=True)
            
        with col2:
            # AGORA VAI FUNCIONAR: as variáveis primeiro_dia_mes e fim_mes agora existem aqui
            st.metric(
                label="Faturamento Mensal", 
                value=f"R$ {fat_mensal:,.2f}", 
                delta=f"{var_mes:+.1f}% (LY: {fat_mensal_ly:,.0f})", 
                help=f"Período: {primeiro_dia_mes.strftime('%d/%m')} até {fim_mes.strftime('%d/%m')}"
            )
            st.plotly_chart(plotar_grafico_comparativo(fat_mensal, fat_mensal_ly), use_container_width=True)
            
        with col3:
            st.metric("Faturamento Anual", f"R$ {fat_anual:,.2f}", f"{var_ano:+.1f}% (LY: {fat_anual_ly:,.0f})")
            st.plotly_chart(plotar_grafico_comparativo(fat_anual, fat_anual_ly), use_container_width=True)

    # -----------------------------------------------------------
    # VISÃO 2: RANKINGS E TABELAS
    # -----------------------------------------------------------
    elif st.session_state['visao_atual'] == 'rankings':
        
        # Filtros para os gráficos (Replicando a lógica de datas)
        filtro_semana = (df['Data de entrega'] >= inicio_sem) & (df['Data de entrega'] <= fim_sem)
        df_semana = df[filtro_semana]

        hoje = pd.Timestamp.now().normalize()
        if hoje.day <= 5:
            inicio_mes = (hoje.replace(day=1) - pd.Timedelta(days=1)).replace(day=1)
            fim_mes = hoje.replace(day=1) - pd.Timedelta(days=1)
        else:
            inicio_mes = hoje.replace(day=1)
            fim_mes = hoje
        
        df_mes = df[(df['Data de entrega'] >= inicio_mes) & (df['Data de entrega'] <= fim_mes)]

        st.subheader(f"🏆 Rankings da Semana ({inicio_sem.strftime('%d/%m')} a {fim_sem.strftime('%d/%m')})")
        
        c_rank1, c_rank2 = st.columns(2)

        # Gráfico 1: Top 5 Faturamento
        with c_rank1:
            st.markdown("#### 💰 Top 5 Faturamento")
            if not df_semana.empty:
                top_fat = df_semana.groupby('LOJAS')['Valor'].sum().reset_index().sort_values('Valor', ascending=False).head(5)
                fig_fat = px.bar(top_fat, x='Valor', y='LOJAS', orientation='h', text_auto='.2s', color_discrete_sequence=['#0068c9'])
                fig_fat.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(l=0,r=0,t=0,b=0), height=300)
                st.plotly_chart(fig_fat, use_container_width=True)
            else:
                st.info("Sem dados de vendas nesta semana ainda.")

        # Gráfico 2: Top 5 Pedidos
        with c_rank2:
            st.markdown("#### 📦 Top 5 Pedidos (Qtd)")
            if not df_semana.empty:
                top_ped = df_semana.groupby('LOJAS')['Valor'].count().reset_index()
                top_ped.columns = ['LOJAS', 'Qtd']
                top_ped = top_ped.sort_values('Qtd', ascending=False).head(5)
                fig_ped = px.bar(top_ped, x='Qtd', y='LOJAS', orientation='h', text_auto=True, color_discrete_sequence=['#29b5e8'])
                fig_ped.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(l=0,r=0,t=0,b=0), height=300)
                st.plotly_chart(fig_ped, use_container_width=True)
            else:
                st.info("Sem pedidos nesta semana ainda.")

        st.markdown("---")
        st.subheader("⚠️ Lojas sem Vendas (Mês Acumulado)")

        todas_lojas = df['LOJAS'].unique()
        lojas_com_venda = df_mes['LOJAS'].unique()
        lojas_zeradas = [loja for loja in todas_lojas if loja not in lojas_com_venda]
        ignorar = ['MIZUNO',
            'MMARTAN',
            'TNG',
            'NÃO IDENTIFICADO',
            'CLUB MARISOL',
            'NEW BALANCE',
            'NIKE',
            'Nike',
            '',
            ' ',
            'PERFUMES OUTLET',
            'ESTOQUE',
            'LUCIANA RANGEL',
            'KOPENHAGEN',
            'OAKLEY',
            'TIP TOP',
            'HAVAIANAS',
            'PLANET GIRLS',
            'INFORTCELL',
            'nan',
            'GUARARAPES',
            'PARAMGABA',] 
        lojas_zeradas = [l for l in lojas_zeradas if str(l).strip() not in ignorar]

        if len(lojas_zeradas) > 0:
            st.error(f"Total de lojas zeradas: {len(lojas_zeradas)}")
            df_zeradas = pd.DataFrame(lojas_zeradas, columns=['LOJAS SEM VENDA'])
            st.dataframe(df_zeradas, hide_index=True, use_container_width=True)
        else:
            st.success("🎉 Parabéns! Todas as lojas venderam neste mês!")

else:
    st.warning("⚠️ Base de dados não encontrada.")