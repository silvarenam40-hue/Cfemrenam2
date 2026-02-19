import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from pathlib import Path

# Configurar o estilo dos gráficos
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (14, 8)
plt.rcParams['font.size'] = 10

# Caminho do arquivo CSV
csv_file = r"C:\Users\renam.antonio\Desktop\CFEM\CFEM_Arrecadacao_2022_2026.csv"

# Verificar se o arquivo existe
if not Path(csv_file).exists():
    print(f"Erro: Arquivo não encontrado em {csv_file}")
    exit()

print("Carregando dados...")
# Ler o arquivo CSV com separador ponto-e-vírgula
try:
    df = pd.read_csv(csv_file, sep=';', encoding='utf-8')
except UnicodeDecodeError:
    try:
        df = pd.read_csv(csv_file, sep=';', encoding='latin-1')
    except UnicodeDecodeError:
        df = pd.read_csv(csv_file, sep=';', encoding='cp1252')

# Converter colunas numéricas (tratar vírgula decimal brasileira)
df['ValorRecolhido'] = df['ValorRecolhido'].astype(str).str.replace('R$', '').str.replace('.', '').str.replace(',', '.').str.strip().astype(float)
df['QuantidadeComercializada'] = df['QuantidadeComercializada'].astype(str).str.replace(',', '.').str.strip().astype(float)

print(f"Total de registros: {len(df)}")
print(f"Colunas: {df.columns.tolist()}")
print("\nPrimeiras linhas:")
print(df.head())

# Criar pasta para salvar os gráficos
output_dir = Path(csv_file).parent / "graficos"
output_dir.mkdir(exist_ok=True)

# ===== GRÁFICO 1: Arrecadação por Ano =====
print("\nGerando gráfico 1: Arrecadação por Ano...")
plt.figure(figsize=(10, 6))
arrecadacao_ano = df.groupby('Ano')['ValorRecolhido'].sum().sort_index()
arrecadacao_ano.plot(kind='bar', color='steelblue', edgecolor='black')
plt.title('Arrecadação CFEM por Ano', fontsize=14, fontweight='bold')
plt.xlabel('Ano', fontsize=12)
plt.ylabel('Valor Recolhido (R$)', fontsize=12)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(output_dir / '01_arrecadacao_por_ano.png', dpi=300, bbox_inches='tight')
plt.close()

# ===== GRÁFICO 2: Top 10 Substâncias mais Arrecadadas =====
print("Gerando gráfico 2: Top 10 Substâncias...")
plt.figure(figsize=(12, 6))
top_substancias = df.groupby('Substância')['ValorRecolhido'].sum().sort_values(ascending=False).head(10)
top_substancias.plot(kind='barh', color='teal', edgecolor='black')
plt.title('Top 10 Substâncias por Valor Arrecadado', fontsize=14, fontweight='bold')
plt.xlabel('Valor Recolhido (R$)', fontsize=12)
plt.ylabel('Substância', fontsize=12)
plt.tight_layout()
plt.savefig(output_dir / '02_top_substancias.png', dpi=300, bbox_inches='tight')
plt.close()

# ===== GRÁFICO 3: Arrecadação por Estado (Top 15) =====
print("Gerando gráfico 3: Arrecadação por Estado...")
plt.figure(figsize=(12, 6))
arrecadacao_uf = df.groupby('UF')['ValorRecolhido'].sum().sort_values(ascending=False).head(15)
arrecadacao_uf.plot(kind='bar', color='coral', edgecolor='black')
plt.title('Top 15 Estados por Arrecadação CFEM', fontsize=14, fontweight='bold')
plt.xlabel('Estado (UF)', fontsize=12)
plt.ylabel('Valor Recolhido (R$)', fontsize=12)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(output_dir / '03_arrecadacao_por_estado.png', dpi=300, bbox_inches='tight')
plt.close()

# ===== GRÁFICO 4: Tendência Mensal de Arrecadação =====
print("Gerando gráfico 4: Tendência Mensal...")
plt.figure(figsize=(14, 6))
# Criar coluna de ano-mês para melhor visualização
df['AnoMes'] = df['Ano'].astype(str) + '-' + df['Mês'].astype(str).str.zfill(2)
arrecadacao_mes = df.groupby('AnoMes')['ValorRecolhido'].sum().sort_index()
plt.plot(range(len(arrecadacao_mes)), arrecadacao_mes.values, marker='o', linewidth=2, color='darkgreen')
plt.fill_between(range(len(arrecadacao_mes)), arrecadacao_mes.values, alpha=0.3, color='lightgreen')
plt.title('Tendência de Arrecadação Mensal', fontsize=14, fontweight='bold')
plt.xlabel('Período', fontsize=12)
plt.ylabel('Valor Recolhido (R$)', fontsize=12)
plt.xticks(range(0, len(arrecadacao_mes), 12), arrecadacao_mes.index[::12], rotation=45)
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig(output_dir / '04_tendencia_mensal.png', dpi=300, bbox_inches='tight')
plt.close()

# ===== GRÁFICO 5: Distribuição por Tipo (PF vs PJ) =====
print("Gerando gráfico 5: Distribuição PF vs PJ...")
plt.figure(figsize=(10, 6))
distribuicao_tipo = df.groupby('Tipo_PF_PJ')['ValorRecolhido'].sum()
cores = ['#ff9999', '#66b3ff']
plt.pie(distribuicao_tipo.values, labels=distribuicao_tipo.index, autopct='%1.1f%%', 
        colors=cores, startangle=90, textprops={'fontsize': 12})
plt.title('Arrecadação: Pessoa Física vs Jurídica', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig(output_dir / '05_distribuicao_pf_pj.png', dpi=300, bbox_inches='tight')
plt.close()

# ===== GRÁFICO 6: Top 10 Municípios =====
print("Gerando gráfico 6: Top 10 Municípios...")
plt.figure(figsize=(12, 6))
top_municipios = df.groupby('Município')['ValorRecolhido'].sum().sort_values(ascending=False).head(10)
top_municipios.plot(kind='barh', color='purple', edgecolor='black')
plt.title('Top 10 Municípios por Arrecadação', fontsize=14, fontweight='bold')
plt.xlabel('Valor Recolhido (R$)', fontsize=12)
plt.ylabel('Município', fontsize=12)
plt.tight_layout()
plt.savefig(output_dir / '06_top_municipios.png', dpi=300, bbox_inches='tight')
plt.close()

# ===== GRÁFICO 7: Arrecadação por Substância e Ano (Top 5) =====
print("Gerando gráfico 7: Substâncias mais importantes por Ano...")
plt.figure(figsize=(14, 6))
top_5_substancias = df.groupby('Substância')['ValorRecolhido'].sum().sort_values(ascending=False).head(5).index
df_top5 = df[df['Substância'].isin(top_5_substancias)]
pivot_data = df_top5.groupby(['Ano', 'Substância'])['ValorRecolhido'].sum().unstack()
pivot_data.plot(kind='bar', ax=plt.gca(), edgecolor='black')
plt.title('Arrecadação das Top 5 Substâncias por Ano', fontsize=14, fontweight='bold')
plt.xlabel('Ano', fontsize=12)
plt.ylabel('Valor Recolhido (R$)', fontsize=12)
plt.legend(title='Substância', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(output_dir / '07_substancias_por_ano.png', dpi=300, bbox_inches='tight')
plt.close()

# ===== ESTATÍSTICAS RESUMIDAS =====
print("\n" + "="*60)
print("ESTATÍSTICAS RESUMIDAS DA ARRECADAÇÃO CFEM")
print("="*60)
print(f"\nPeríodo: {df['Ano'].min()} a {df['Ano'].max()}")
print(f"Total de registros: {len(df):,}")
print(f"Total arrecadado: R$ {df['ValorRecolhido'].sum():,.2f}")
print(f"Arrecadação média por registro: R$ {df['ValorRecolhido'].mean():,.2f}")
print(f"\nTop 5 Substâncias:")
print(df.groupby('Substância')['ValorRecolhido'].sum().sort_values(ascending=False).head(5))
print(f"\nTop 5 Estados:")
print(df.groupby('UF')['ValorRecolhido'].sum().sort_values(ascending=False).head(5))
print(f"\nDistribuição PF/PJ:")
print(df.groupby('Tipo_PF_PJ')['ValorRecolhido'].sum())

print(f"\n✅ Análise concluída! Gráficos salvos em: {output_dir}")
print("="*60)
