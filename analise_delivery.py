import pandas as pd
import os
import pywhatkit
import time
from dotenv import load_dotenv
#=============================================================================
# 1 - CARREGAR DADOS E LIMPEZA DE DADOS
#=============================================================================

df = pd.read_excel(r"G:/Meu Drive/Delivery/Delivery/delivery_atualizado.xlsx", sheet_name='Dados Gerais')

# Remove espaços vazios antes e depois dos nomes das colunas (ex: 'Valor ' vira 'Valor')
df['LOJAS'] = df['LOJAS'].str.strip().str.upper()
df.columns = df.columns.str.strip()

# 3. Converter Colunas para o formato certo
# Garante que Data de entrega é data. 'errors=coerce' transforma erros em NaNs (vazio) para não travar
df['Data de entrega'] = pd.to_datetime(df['Data de entrega'], errors='coerce')

# Garante que Valor é número. 'errors=coerce' transforma erros em NaNs (vazio) para não travar
df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce')

print(df.head())

#=============================================================================
# 2 - SETANDO PERÍODOS DE ANÁLISE E IMPRIMINDO RESULTADOS - SEMANA
#=============================================================================

hoje = pd.Timestamp.now().normalize()

idx_domingo = (hoje.weekday() + 1) % 7
ultimo_domingo = hoje - pd.Timedelta(days=idx_domingo)

fim_sem_completa = ultimo_domingo
inicio_sem_completa = fim_sem_completa - pd.Timedelta(days=6)

inicio_semana_ly = inicio_sem_completa - pd.Timedelta(days=364)
fim_semana_ly = fim_sem_completa - pd.Timedelta(days=364)
filtro_atual = (df['Data de entrega'] >= inicio_sem_completa) & (df['Data de entrega'] <= fim_sem_completa)
filtro_ly = (df['Data de entrega'] >= inicio_semana_ly) & (df['Data de entrega'] <= fim_semana_ly)

df_atual = df[filtro_atual]
df_ly = df[filtro_ly]


faturamento_atual = df_atual['Valor'].sum()
faturamento_ly = df_ly['Valor'].sum()
entregas_atual = df_atual.shape[0]
entregas_ly = df_ly.shape[0]

print(f"--- RELATÓRIO AUTOMÁTICO ({hoje.date()}) ---")
print(f"Período Atual: {inicio_sem_completa.date()} até {fim_sem_completa.date()}")
print(f"Período LY:    {inicio_semana_ly.date()} até {fim_semana_ly.date()}\n")

print(f"💰 Faturamento Atual: R$ {faturamento_atual:,.2f}")
print(f"💰 Faturamento LY:    R$ {faturamento_ly:,.2f}")


if faturamento_ly > 0:
    var_fat = ((faturamento_atual - faturamento_ly) / faturamento_ly) * 100
    print(f"   Variação: {var_fat:+.2f}%") # O "+" força mostrar sinal de positivo ou negativo
else:
    print("   Variação: - (Sem dados no ano anterior)")

print(f"\n📦 Entregas Atual: {entregas_atual}")
print(f"📦 Entregas LY:    {entregas_ly}")

if entregas_ly > 0:
    var_ent = ((entregas_atual - entregas_ly) / entregas_ly) * 100
    print(f"   Variação: {var_ent:+.2f}%")
else:
    print("   Variação: - (Sem dados no ano anterior)")


#=============================================================================
# 3 - SETANDO PERÍODOS DE ANÁLISE E IMPRIMINDO RESULTADOS - MES
#=============================================================================

# ---- Visão Mensal (MTD) ---- #
# Define o período MTD (Month-to-Date)

if hoje.day <= 5:
    FECHAMENTO_MES_ANTERIOR = True
    print(f"🤖 Dia {hoje.day}: Modo FECHAMENTO ativado automaticamente.")
else:
    FECHAMENTO_MES_ANTERIOR = False
    print(f"🤖 Dia {hoje.day}: Modo ACOMPANHAMENTO ativado automaticamente.")

if FECHAMENTO_MES_ANTERIOR:
    # Lógica: Queremos o mês passado inteiro
    # Ex: Se hoje é 02/02, queremos de 01/01 até 31/01
    
    primeiro_dia_deste_mes = hoje.replace(day=1)
    # O fim do período é o dia anterior ao dia 01 deste mês (ou seja, 31 do mês passado)
    fim_mes_atual = primeiro_dia_deste_mes - pd.Timedelta(days=1)
    # O início é o dia 1 daquele mês
    inicio_mes_atual = fim_mes_atual.replace(day=1)
    
    texto_periodo = "FECHAMENTO MÊS ANTERIOR"
    
else:
    # Lógica Original: Mês Atual até hoje (MTD)
    inicio_mes_atual = hoje.replace(day=1)
    fim_mes_atual = hoje
    texto_periodo = "PARCIAL MÊS ATUAL (MTD)"


inicio_mes_ly    = inicio_mes_atual - pd.DateOffset(years=1) # Dia 1 do mesmo mês ano passado
fim_mes_ly       = fim_mes_atual - pd.DateOffset(years=1)

filtro_mtd_atual = (df['Data de entrega'] >= inicio_mes_atual) & (df['Data de entrega'] <= fim_mes_atual)
filtro_mtd_ly    = (df['Data de entrega'] >= inicio_mes_ly) & (df['Data de entrega'] <= fim_mes_ly)
df_mtd_atual = df[filtro_mtd_atual]
df_mtd_ly    = df[filtro_mtd_ly]

#---- Acumulado do mês (MTD) ----#
print(f"\n🗓️  VISÃO MENSAL: {texto_periodo}")
print(f"   Periodo: {inicio_mes_atual.strftime('%d/%m')} até {fim_mes_atual.strftime('%d/%m')}")

fat_mtd = df_mtd_atual['Valor'].sum()
fat_mtd_ly = df_mtd_ly['Valor'].sum()
var_mtd = ((fat_mtd - fat_mtd_ly) / fat_mtd_ly * 100) if fat_mtd_ly > 0 else 0

entregas_mtd = len(df_mtd_atual)
entregas_mtd_ly = len(df_mtd_ly)
var_entregas_mtd = ((entregas_mtd - entregas_mtd_ly) / entregas_mtd_ly * 100) if entregas_mtd_ly > 0 else 0

print(f"   💰 Faturamento: R$ {fat_mtd:,.2f} (LY: R$ {fat_mtd_ly:,.2f})")
print(f"      Crescimento: {var_mtd:+.1f}%")
print(f"   📦 Pedidos:     {entregas_mtd} (LY: {entregas_mtd_ly})")
print(f"      Crescimento: {var_entregas_mtd:+.1f}%")


#=============================================================================
# 4 - RANKING DE LOJAS
#============================================================================

# Agrupamos por LOJA e pedimos duas contas:
# - Coluna 'Valor': queremos a SOMA ('sum') -> Faturamento
# - Coluna 'LOJAS': queremos a CONTAGEM ('count') -> Quantidade de Pedidos
ranking = df_atual.groupby('LOJAS').agg({
    'Valor': 'sum',
    'LOJAS': 'count'
})

ranking = ranking.rename(columns={
    'Valor': 'Faturamento',
    'LOJAS': 'Quantidade de Pedidos'
})

ranking_mensal = df_mtd_atual.groupby('LOJAS').agg({
    'Valor': 'sum',
    'LOJAS': 'count'
})
ranking_mensal = ranking_mensal.rename(columns={
    'Valor': 'Faturamento',
    'LOJAS': 'Quantidade de Pedidos'
})

todas_lojas = df['LOJAS'].unique()
ranking = ranking.reindex(todas_lojas, fill_value=0)
ranking_mensal = ranking_mensal.reindex(todas_lojas, fill_value=0)


#==============================================================================
# 5 - IMPRESSÃO DO RANKING E LOJAS ZERADAS
#==============================================================================

lojas_para_ignorar = [
    'MIZUNO',
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
    'PARAMGABA',
]

ranking = ranking.drop(lojas_para_ignorar, errors='ignore')
ranking_mensal = ranking_mensal.drop(lojas_para_ignorar, errors='ignore')

lojas_zeradas = ranking_mensal[ranking_mensal['Quantidade de Pedidos'] == 0]
ranking_com_vendas = ranking[ranking['Faturamento'] > 0]

top_faturamento = ranking.sort_values(by='Faturamento', ascending=False).head(5)
top_pedidos = ranking.sort_values(by='Quantidade de Pedidos', ascending=False).head(5)
menor_faturamento = ranking_com_vendas.sort_values(by='Faturamento', ascending=True)
menor_pedidos = ranking_com_vendas.sort_values(by='Quantidade de Pedidos', ascending=True)



# iterrows() permite passar linha por linha para formatar bonito
print("\n📦 TOP 5 - MAIOR VOLUME DE FATURAMENTO")
for loja, dados in top_faturamento.iterrows():
    print(f"   {loja:<20} | R$ {dados['Faturamento']:,.2f} ({int(dados['Quantidade de Pedidos'])} peds)")

print("\n📦 TOP 5 - MAIOR VOLUME DE PEDIDOS")
for loja, dados in top_pedidos.iterrows():
    print(f"   {loja:<20} | {int(dados['Quantidade de Pedidos'])} pedidos (R$ {dados['Faturamento']:,.2f})")

print("\n📦 TOP 20 - MENOR VOLUME DE PEDIDOS")
for loja, dados in menor_faturamento.iterrows():
    print(f"   {loja:<20} | R$ {dados['Faturamento']:,.2f} ({int(dados['Quantidade de Pedidos'])} peds)")
    

print("\n📦 TOP 20 - MENOR FATURAMENTO")
for loja, dados in menor_pedidos.iterrows():
    print(f"   {loja:<20} | {int(dados['Quantidade de Pedidos'])} pedidos (R$ {dados['Faturamento']:,.2f})")

# Se não houver dados na semana (ex: segunda de manhã), avisa
if df_mtd_atual.empty:
    print("\n⚠️ Nenhuma entrega registrada nos últimos 7 dias ainda.")

print(f"\n🚨 LOJAS SEM MOVIMENTO ({len(lojas_zeradas)} lojas)")
if not lojas_zeradas.empty:
    # Como sabemos que é tudo zero, não precisamos gastar processamento lendo os dados da linha
    for loja in lojas_zeradas.index:
        print(f"   {loja:<20} | R$ 0.00 (0 peds)")
        
print("\n--- FIM DO RELATÓRIO ---")


#=============================================================================
# 6. SALVAR RELATÓRIO DETALHADO EM EXCEL
#=============================================================================

# 1. Pega a pasta onde está o arquivo original
caminho_arquivo = r"G:\Meu Drive\Delivery\Delivery\delivery_atualizado.xlsx"
pasta_original = os.path.dirname(caminho_arquivo)

# 2. Define o nome da pasta nova (pode mudar "Relatorios" para o que quiser)
nome_subpasta = "Relatorios_Gerados"
caminho_da_pasta_nova = os.path.join(pasta_original, nome_subpasta)

# 3. O PULO DO GATO: Cria a pasta se ela não existir
if not os.path.exists(caminho_da_pasta_nova):
    os.makedirs(caminho_da_pasta_nova)
    print(f"📁 Pasta criada automaticamente: {caminho_da_pasta_nova}")

# 4. Define o caminho final do arquivo (dentro da pasta nova)
nome_arquivo = f"Relatorio_Delivery_{hoje.strftime('%d-%m-%Y')}.xlsx"
caminho_completo_saida = os.path.join(caminho_da_pasta_nova, nome_arquivo)

try:
    with pd.ExcelWriter(caminho_completo_saida) as writer:
        
        # Aba 1: Top Faturamento
        top_faturamento.to_excel(writer, sheet_name='Top 5 Faturamento')
        
        # Aba 2: Top Pedidos
        top_pedidos.to_excel(writer, sheet_name='Top 5 Pedidos')
        
        # Aba 3: Menor Faturamento (Quem vendeu pouco)
        menor_faturamento.to_excel(writer, sheet_name='Menor Faturamento')
        
        # Aba 4: Lojas Zeradas (Se houver)
        if not lojas_zeradas.empty:
            lojas_zeradas.to_excel(writer, sheet_name='Lojas Zeradas')
            
        # Aba 5: Resumo Geral
        resumo = pd.DataFrame({
            'Indicador': [
                'ATUAL SEMANA: FATURAMENTO', 
                'LY SEMANA: FATURAMENTO',
                'CRESCIMENTO FATURAMENTO',
                '---',
                'ATUAL SEMANA: PEDIDOS',
                'LY SEMANA: PEDIDOS',
                'CRESCIMENTO PEDIDOS',
                '---', 
                'Lojas Zeradas',
                '---',  # Uma linha separadora visual
                'ATUAL MÊS: FATURAMENTO ACUMULADO: ',
                'LY MES: FATURAMENTO ACUMULADO: ',
                'CRESCIMENTO FATURAMENTO',
                '---',
                'ATUAL MES: PEDIDOS ACUMULADO',
                'LY MES: PEDIDOS ACUMULADO',
                'CRESCIMENTO PEDIDOS',
            ],
            'Valor': [
                faturamento_atual,
                faturamento_ly,
                f"{var_fat:+.2f}%", # Formata como porcentagem bonita
                '', # Valor vazio para o separador 
                entregas_atual,
                entregas_ly,
                f"{var_ent:+.2f}%", # Formata como porcentagem bonita
                '', # Valor vazio para o separador 
                len(lojas_zeradas),
                '', # Valor vazio para o separador
                fat_mtd,
                fat_mtd_ly,
                f"{var_mtd:+.1f}%", # Formata como porcentagem bonita
                '', # Valor vazio para o separador
                entregas_mtd,
                entregas_mtd_ly,
                f"{var_entregas_mtd:+.1f}%", # Formata como porcentagem bonita
            ]
        })
        resumo.to_excel(writer, sheet_name='Resumo Geral', index=False)

    print(f"✅ RELATÓRIO SALVO COM SUCESSO!")
    print(f"📂 Local: {caminho_completo_saida}")

except PermissionError:
    print("❌ ERRO: O arquivo Excel já está aberto!")
    print("Feche o arquivo 'Relatorio_Delivery...' e tente rodar novamente.")
except Exception as e:
    print(f"❌ Erro inesperado ao salvar: {e}")


# ==============================================================================
# 7. ENVIO DE RESUMO GERENCIAL NO WHATSAPP
# ==============================================================================

print("\n📱 Preparando mensagem do WhatsApp...")

# --- Configuração do Título Dinâmico ---
# Aqui ele verifica qual opção você escolheu lá em cima
if 'FECHAMENTO_MES_ANTERIOR' in locals() and FECHAMENTO_MES_ANTERIOR:
    titulo_mes = "🗓️ FECHAMENTO MÊS ANTERIOR"
    # Formata: "01/01 a 31/01"
    periodo_txt = f"({inicio_mes_atual.strftime('%d/%m')} a {fim_mes_atual.strftime('%d/%m')})"
else:
    titulo_mes = "🗓️ ACUMULADO MÊS (MTD)"
    periodo_txt = "(Mês Atual)"

MEU_NUMERO = "+5585999210476" # <--- SEU NÚMERO AQUI

mensagem_zap = f"""
*🤖 RELATÓRIO DELIVERY*

*🗓️ SEMANA ATUAL*
💰 Fat: R$ {faturamento_atual:,.2f} ({var_fat:+.1f}%)
📦 Ped: {entregas_atual} ({var_ent:+.1f}%)

*{titulo_mes}*
_{periodo_txt}_
💰 Fat: R$ {fat_mtd:,.2f} ({var_mtd:+.1f}%)
📦 Ped: {entregas_mtd} ({var_entregas_mtd:+.1f}%)

*📉 ALERTAS*
Lojas Zeradas: {len(lojas_zeradas)}
_Top 1 Fat: {top_faturamento.index[0]} (R$ {top_faturamento.iloc[0]['Faturamento']:,.0f})_

_📁 Relatório detalhado salvo em:_
_{nome_arquivo}_
"""

try:
    pywhatkit.sendwhatmsg_instantly(
        phone_no=os.getenv("WHATSAPP_NUMERO"), 
        message=mensagem_zap, 
        wait_time=20, 
        tab_close=True
    )
    print("✅ Mensagem enviada com sucesso!")
    
except Exception as e:
    print(f"❌ Erro no Zap: {e}")