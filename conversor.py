import pandas as pd
import json
from datetime import datetime

arquivo = "T20 - Tabela de geração de tesouros 2.0.xlsx"
xls = pd.ExcelFile(arquivo)
bd = {}

def arrumar_chance(valor):
    if pd.isna(valor) or valor == '—' or valor == 'd%': return None
    # Corrige datas (o Excel converte chances como "01-10" para 1 de Outubro)
    if isinstance(valor, pd.Timestamp) or isinstance(valor, datetime): 
        return f"{valor.day:02d}-{valor.month:02d}"
    val_str = str(valor).strip()
    if '-' in val_str: return val_str
    try: return f"{int(val_str):02d}-{int(val_str):02d}"
    except: return val_str

def extrair_tabela(nome_aba, col_nome):
    if nome_aba not in xls.sheet_names: return []
    df = pd.read_excel(xls, sheet_name=nome_aba)
    
    col_espaco = None
    for col in df.columns:
        if "espaço" in str(col).lower() or "espaco" in str(col).lower():
            col_espaco = col
            break

    lista = []
    for _, row in df.iterrows():
        chance = arrumar_chance(row.get('d%'))
        if chance and pd.notna(row.get(col_nome)):
            livro = str(row.get('Livro', ''))
            pagina = str(row.get('Página', ''))
            espacos = str(row[col_espaco]) if col_espaco and pd.notna(row[col_espaco]) else ""
            
            if livro == 'nan': livro = ""
            if pagina == 'nan': pagina = ""
            if espacos == 'nan' or espacos == 'None': espacos = ""
            
            lista.append({
                "chance": chance, "nome": str(row[col_nome]), 
                "livro": livro, "pagina": pagina, "espacos": espacos
            })
    return lista

def extrair_multi_tabela(nome_aba):
    df = pd.read_excel(xls, sheet_name=nome_aba)
    tabela = {"Armas": [], "Armaduras": [], "Esotericos": []}
    for index, row in df.iterrows():
        if index == 0: continue
        
        c_arma = arrumar_chance(row.iloc[0])
        if c_arma and pd.notna(row.iloc[1]): 
            tabela['Armas'].append({"chance": c_arma, "nome": str(row.iloc[1]), "livro": str(row.iloc[2]) if pd.notna(row.iloc[2]) else "", "pagina": str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""})
        
        c_armad = arrumar_chance(row.iloc[5])
        if c_armad and pd.notna(row.iloc[6]): 
            tabela['Armaduras'].append({"chance": c_armad, "nome": str(row.iloc[6]), "livro": str(row.iloc[7]) if pd.notna(row.iloc[7]) else "", "pagina": str(row.iloc[8]) if pd.notna(row.iloc[8]) else ""})
        
        c_esot = arrumar_chance(row.iloc[10])
        if c_esot and pd.notna(row.iloc[11]): 
            tabela['Esotericos'].append({"chance": c_esot, "nome": str(row.iloc[11]), "livro": str(row.iloc[12]) if pd.notna(row.iloc[12]) else "", "pagina": str(row.iloc[13]) if pd.notna(row.iloc[13]) else ""})
    return tabela

def extrair_riquezas():
    df = pd.read_excel(xls, sheet_name='Riquezas')
    tabela = {"menor": [], "media": [], "maior": []}

    for i_row, row in df.iterrows():
        # Pula as linhas de cabeçalho perdidas pelo Excel
        if "MENOR" in str(row.iloc[0]).upper() or "D%" in str(row.iloc[0]).upper() or "RIQUEZA" in str(row.iloc[0]).upper():
            continue

        # Garante que a linha tem colunas suficientes e possui texto de exemplo
        if len(row) < 5 or pd.isna(row.iloc[4]):
            continue

        # O texto da riqueza está na coluna 4 (Exemplos) e o valor está na coluna 3 (Valor T$)
        texto_exemplo = str(row.iloc[4]).strip()
        valor_ts = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
        
        # Junta o valor em dinheiro com os exemplos para exibir tudo no card
        nome_final = f"<span style='color:#fbc02d;'>[T$ {valor_ts}]</span> {texto_exemplo}" if valor_ts and valor_ts != "nan" else texto_exemplo

        # Menor (Coluna 0)
        c_men = arrumar_chance(row.iloc[0])
        if c_men and str(c_men) not in ['-', '–', '—', 'nan']:
            tabela["menor"].append({"chance": c_men, "nome": nome_final, "espacos": ""})

        # Média (Coluna 1)
        c_med = arrumar_chance(row.iloc[1])
        if c_med and str(c_med) not in ['-', '–', '—', 'nan']:
            tabela["media"].append({"chance": c_med, "nome": nome_final, "espacos": ""})

        # Maior (Coluna 2)
        c_mai = arrumar_chance(row.iloc[2])
        if c_mai and str(c_mai) not in ['-', '–', '—', 'nan']:
            tabela["maior"].append({"chance": c_mai, "nome": nome_final, "espacos": ""})

    return tabela

def extrair_acessorios():
    df = pd.read_excel(xls, sheet_name='Mágicos (Acessórios)')
    tabela = {"Menor": [], "Medio": [], "Maior": []}
    for index, row in df.iterrows():
        if index == 0: continue
        
        c_men = arrumar_chance(row.iloc[0])
        if c_men and pd.notna(row.iloc[1]): 
            tabela['Menor'].append({"chance": c_men, "nome": str(row.iloc[1]), "livro": str(row.iloc[3]) if pd.notna(row.iloc[3]) else "", "pagina": str(row.iloc[4]) if pd.notna(row.iloc[4]) else ""})
        
        c_med = arrumar_chance(row.iloc[6])
        if c_med and pd.notna(row.iloc[7]): 
            tabela['Medio'].append({"chance": c_med, "nome": str(row.iloc[7]), "livro": str(row.iloc[9]) if pd.notna(row.iloc[9]) else "", "pagina": str(row.iloc[10]) if pd.notna(row.iloc[10]) else ""})
        
        c_mai = arrumar_chance(row.iloc[12])
        if c_mai and pd.notna(row.iloc[13]): 
            tabela['Maior'].append({"chance": c_mai, "nome": str(row.iloc[13]), "livro": str(row.iloc[15]) if pd.notna(row.iloc[15]) else "", "pagina": str(row.iloc[16]) if pd.notna(row.iloc[16]) else ""})
    return tabela

def extrair_magicos():
    df = pd.read_excel(xls, sheet_name='Mágicos')
    tabela = {"Armas": [], "Armaduras": [], "Esotericos": [], "Armas_Esp": [], "Armaduras_Esp": [], "Esotericos_Esp": []}
    
    modo_arma = "Armas"
    modo_armad = "Armaduras"
    modo_esot = "Esotericos"
    
    for index, row in df.iterrows():
        # Verifica as colunas procurando pelos cabeçalhos pretos no PLURAL que dividem as tabelas
        str_arma = str(row.iloc[0]).upper() + " " + str(row.iloc[1]).upper()
        if "ESPECÍFICAS" in str_arma or "ESPECÍFICOS" in str_arma: modo_arma = "Armas_Esp"
        
        str_armad = str(row.iloc[5]).upper() + " " + str(row.iloc[6]).upper()
        if "ESPECÍFICAS" in str_armad or "ESPECÍFICOS" in str_armad: modo_armad = "Armaduras_Esp"
        
        str_esot = str(row.iloc[10]).upper() + " " + str(row.iloc[11]).upper()
        if "ESPECÍFICAS" in str_esot or "ESPECÍFICOS" in str_esot: modo_esot = "Esotericos_Esp"

        # Armas
        c_arma = arrumar_chance(row.iloc[0])
        if c_arma and pd.notna(row.iloc[1]) and str(row.iloc[1]).lower() != "encanto":
            tabela[modo_arma].append({"chance": c_arma, "nome": str(row.iloc[1]), "livro": str(row.iloc[2]) if pd.notna(row.iloc[2]) else "", "pagina": str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""})
            
        # Armaduras
        c_armad = arrumar_chance(row.iloc[5])
        if c_armad and pd.notna(row.iloc[6]) and str(row.iloc[6]).lower() != "encanto":
            tabela[modo_armad].append({"chance": c_armad, "nome": str(row.iloc[6]), "livro": str(row.iloc[7]) if pd.notna(row.iloc[7]) else "", "pagina": str(row.iloc[8]) if pd.notna(row.iloc[8]) else ""})
            
        # Esotéricos
        c_esot = arrumar_chance(row.iloc[10])
        if c_esot and pd.notna(row.iloc[11]) and str(row.iloc[11]).lower() != "encanto":
            tabela[modo_esot].append({"chance": c_esot, "nome": str(row.iloc[11]), "livro": str(row.iloc[12]) if pd.notna(row.iloc[12]) else "", "pagina": str(row.iloc[13]) if pd.notna(row.iloc[13]) else ""})
            
    return tabela

# 1. Tesouro por ND
df_nd = pd.read_excel(xls, sheet_name='Tesouro por ND')
df_nd['ND'] = df_nd['ND'].ffill()
def arrumar_nd(v):
    # Se o Excel converteu a fração para uma data
    if isinstance(v, pd.Timestamp) or hasattr(v, 'month'):
        if v.month == 4 or v.day == 4: return '1/4'
        if v.month == 2 or v.day == 2: return '1/2'
        
    v_str = str(v).strip()
    # Se ele leu como um texto de data ou número decimal
    if '-04' in v_str or v_str == '0.25': return '1/4'
    if '-02' in v_str or v_str == '0.5': return '1/2'
    
    return v_str.replace('.0', '')
df_nd['ND'] = df_nd['ND'].apply(arrumar_nd)

bd['tesouro_nd'] = []
for _, row in df_nd.iterrows():
    if pd.isna(row['d%']) and pd.isna(row.get('d%.1')): continue
    bd['tesouro_nd'].append({
        "nd": row['ND'], "chance_dinheiro": arrumar_chance(row['d%']),
        "dinheiro": row['Dinheiro'] if pd.notna(row['Dinheiro']) else "—",
        "chance_itens": arrumar_chance(row.get('d%.1')),
        "itens": row['Itens'] if pd.notna(row['Itens']) else "—"
    })

# 2. Demais Tabelas
bd['itens_diversos'] = extrair_tabela('Itens Diversos', 'Item')
bd['pocoes'] = extrair_tabela('Poções', 'Poção')
bd['riquezas'] = extrair_riquezas()
bd['equipamentos'] = extrair_multi_tabela('Equipamentos')
bd['superiores'] = extrair_multi_tabela('Superiores')
bd['magicos'] = extrair_magicos()
bd['acessorios'] = extrair_acessorios()

with open('dados.json', 'w', encoding='utf-8') as f:
    json.dump(bd, f, ensure_ascii=False, indent=4)
print("Sucesso: 'dados.json' atualizado com todas as tabelas!")