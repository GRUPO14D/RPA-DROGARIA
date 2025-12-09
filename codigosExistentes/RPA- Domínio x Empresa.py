import pandas as pd
import os
import warnings
import re
import sys
import subprocess
from pathlib import Path
from typing import Callable, Optional

# ==============================================================================
# CONFIGURAÇÕES
# ==============================================================================
# Mes/ano padrao caso nao seja informado (MM-AAAA)
MES_ANO_DEFAULT = "11-2025"
TETO_VALOR_ACEITAVEL = 5000000.00 

LISTA_EMPRESAS = [
    "DROGARIA LIMEIRA",
    "DROGARIA MORELLI FILIAL",
    "DROGARIA MORELLI MTZ"
]

BASES_TEMPLATE = [
    r"N:\Matriz-Jds\Arquivos NF-e\{ano}\{mes_ano}",
    r"N:\Arquivos NF-e\{ano}\{mes_ano}",
    r"N:\{ano}\{mes_ano}"
]

DIR_SAIDA_RPA = r"V:\Fiscal\RPA"
KEYWORD_DOMINIO = "DOMINIO"
KEYWORD_EMPRESA = "EMPRESA"
LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\scalc.exe"

warnings.filterwarnings('ignore')

# Logger simples: por padrao usa print, mas pode receber callback (ex.: UI)
LOG_FN: Optional[Callable[[str], None]] = None

def set_logger(fn: Callable[[str], None]):
    """Define callback para registrar mensagens (ex.: self.add_log_message)."""
    global LOG_FN
    LOG_FN = fn

def log(msg: str):
    """Registra mensagem no callback (se houver) e no stdout."""
    if LOG_FN:
        try:
            LOG_FN(msg)
        except Exception:
            # Nao interrompe o fluxo em caso de erro no callback
            pass
    print(msg)

# Auxiliar para preencher celulas de ranges mesclados que ficam vazias apos leitura
def preencher_mesclados(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    try:
        return df.ffill()
    except Exception:
        return df

def converter_para_xlsx(caminho_arquivo: str) -> Optional[Path]:
    """Converte .xls para .xlsx via LibreOffice (headless). Retorna caminho convertido ou None."""
    src = Path(caminho_arquivo)
    if src.suffix.lower() != ".xls":
        return src
    if not Path(LIBREOFFICE_PATH).exists():
        log(f"[ERRO] LibreOffice não encontrado em {LIBREOFFICE_PATH}")
        return None
    dst = src.with_suffix("")  # remove .xls
    dst = dst.with_name(dst.name + "_conv.xlsx")
    cmd = [
        LIBREOFFICE_PATH,
        "--headless",
        "--convert-to", "xlsx",
        "--outdir", str(src.parent),
        str(src)
    ]
    log(f"       -> Convertendo '{src.name}' para XLSX...")
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if res.returncode != 0:
            log(f"[ERRO CONVERSAO] {res.stderr.strip() or res.stdout.strip()}")
            return None
        if not dst.exists():
            log(f"[ERRO CONVERSAO] Arquivo convertido não encontrado: {dst}")
            # tenta procurar qualquer *_conv.xlsx criado na pasta
            candidatos = list(src.parent.glob(f"{src.stem}*.xlsx"))
            if candidatos:
                log(f"-> Usando candidato encontrado: {candidatos[0].name}")
                return candidatos[0]
            return None
        return dst
    except Exception as exc:
        log(f"[ERRO CONVERSAO] {exc}")
        return None

# ==============================================================================
# 1. FUNÇÕES DE LIMPEZA E VALIDAÇÃO
# ==============================================================================

def converter_para_float(texto):
    if pd.isna(texto) or str(texto).strip() == "": return 0.0
    if isinstance(texto, (int, float)): return float(texto)
    t = str(texto).strip()
    t = re.sub(r'[^\d.,-]', '', t)
    if not t: return 0.0
    if ',' in t and '.' in t:
        if t.find(',') > t.find('.'): t = t.replace('.', '').replace(',', '.')
        else: t = t.replace(',', '')
    elif ',' in t: t = t.replace(',', '.')
    try: 
        val = float(t)
        if val > TETO_VALOR_ACEITAVEL: return 0.0
        return val
    except: return 0.0

def normalizar_nota(nota):
    if pd.isna(nota): return "S/N"
    try: 
        so_numeros = re.sub(r'[^\d.]', '', str(nota))
        if not so_numeros: return "S/N"
        val = float(so_numeros)
        if val <= 0: return "S/N"
        return str(int(val))
    except: return str(nota).strip()

def coluna_parece_indice(serie):
    """
    Detecta se a coluna é apenas um contador de linhas (1, 2, 3...)
    para não confundir com o número da Nota Fiscal.
    """
    try:
        # Pega os primeiros 15 valores numéricos
        vals = pd.to_numeric(serie, errors='coerce').dropna().head(15).tolist()
        if len(vals) < 5: return False
        
        # Verifica se são pequenos (notas fiscais geralmente são > 100)
        media = sum(vals) / len(vals)
        if media < 50: 
            return True # É índice (1, 2, 3...)
            
        # Verifica se é perfeitamente sequencial
        diferencas = [vals[i+1] - vals[i] for i in range(len(vals)-1)]
        if all(d == 1 for d in diferencas):
            return True # É índice
            
        return False
    except:
        return False

# ==============================================================================
# 2. LEITURA
# ==============================================================================

def ler_arquivo_na_rede(caminho_arquivo):
    nome_arq = os.path.basename(caminho_arquivo)
    log(f"    ...Lendo: {nome_arq}")

    caminho_para_ler = converter_para_xlsx(caminho_arquivo)
    if not caminho_para_ler:
        return None

    try:
        df = pd.read_excel(caminho_para_ler, header=None, engine="openpyxl")
        return preencher_mesclados(df)
    except Exception as exc:
        log(f"       [ERRO LEITURA] {exc}")
        return None
def localizar_inicio_tabela(df):
    if df is None or df.empty: return df
    log("       -> Varrendo cabeçalho...")
    for i in range(min(30, len(df))):
        linha_texto = " ".join(df.iloc[i].astype(str).str.lower().tolist())
        # Critério: Tem que ter "nota" E ("valor" ou "total")
        if ("nota" in linha_texto or "n.º" in linha_texto) and ("valor" in linha_texto or "total" in linha_texto):
            log(f"          [OK] Cabeçalho na linha {i+1}")
            df.columns = df.iloc[i].astype(str).str.strip()
            return df[i+1:].copy()
    return df

def preparar_dataframe(df_raw, tipo_origem):
    """
    Prepara dataframe com mapeamento fixo:
    - DOMINIO: Nota = coluna E (indice 4); Valor Contabil = coluna W (indice 22)
    - EMPRESA: N.Nota = coluna M (indice 12); Total Nota = coluna R (indice 17)
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame(columns=["Nota", "Valor"])

    # Se veio do Regex manual (lista de numeros), tenta localizar cabecalho
    if isinstance(df_raw.columns[0], int):
        df_raw = localizar_inicio_tabela(df_raw)

    df_raw.columns = df_raw.columns.astype(str).str.lower().str.strip()

    if tipo_origem == "DOMINIO":
        if len(df_raw.columns) <= 22:
            log("[ERRO] Relatorio DOMINIO com colunas insuficientes")
            return pd.DataFrame(columns=["Codigo", "Nota", "Valor"])
        col_nota = df_raw.columns[4]
        col_valor = df_raw.columns[22]
        possivel_cod = [c for c in df_raw.columns if "codigo" in c or "c?digo" in c]
        col_cod = possivel_cod[0] if possivel_cod else None
        log(f"          [MAPA FIXO DOMINIO] Nota: '{col_nota}' | Valor: '{col_valor}'")
    else:
        if len(df_raw.columns) <= 17:
            log("[ERRO] Relatorio EMPRESA com colunas insuficientes")
            return pd.DataFrame(columns=["Codigo", "Nota", "Valor"])
        col_nota = df_raw.columns[12]
        col_valor = df_raw.columns[17]
        possivel_cod = [c for c in df_raw.columns if "codigo" in c or "c?digo" in c]
        col_cod = possivel_cod[0] if possivel_cod else None
        log(f"          [MAPA FIXO EMPRESA] Nota: '{col_nota}' | Valor: '{col_valor}'")

    try:
        cols = [col_nota, col_valor]
        if col_cod:
            cols.insert(0, col_cod)
        df_new = df_raw[cols].copy()
        if col_cod:
            df_new.columns = ["Codigo", "Nota", "Valor"]
        else:
            df_new.columns = ["Nota", "Valor"]
            df_new["Codigo"] = ""
    except Exception as exc:
        log(f"[ERRO] Falha ao recortar colunas: {exc}")
        return pd.DataFrame(columns=["Codigo", "Nota", "Valor"])

    df_new["Nota"] = df_new["Nota"].apply(normalizar_nota)
    df_new["Valor"] = df_new["Valor"].apply(converter_para_float)

    df_new = df_new[
        (df_new["Valor"] > 0.01) &
        (df_new["Valor"] < TETO_VALOR_ACEITAVEL) &
        (df_new["Nota"] != "S/N") &
        (df_new["Nota"] != "0")
    ]

    if not df_new.empty:
        log(f"          -> Amostra Notas: {df_new['Nota'].head(3).tolist()}")

    return df_new.groupby("Nota", as_index=False).agg({"Valor": "sum", "Codigo": "first"})
# ==============================================================================
# 4. PROCESSAMENTO FINAL
# ==============================================================================

def processar_empresa(empresa, pasta_base, mes_ano):
    log(f"\n==================================================")
    log(f"EMPRESA: {empresa}")
    pasta_relatorio = f"RELATORIO RPA - {empresa}"
    path_rpa = os.path.join(pasta_base, empresa, "ESCRITA FISCAL", pasta_relatorio)
    
    f_dom = f_emp = None
    if os.path.exists(path_rpa):
        for f in os.listdir(path_rpa):
            if KEYWORD_DOMINIO in f.upper() and f.endswith(('.xls', '.xlsx')): f_dom = os.path.join(path_rpa, f)
            if KEYWORD_EMPRESA in f.upper() and f.endswith(('.xls', '.xlsx')): f_emp = os.path.join(path_rpa, f)

    if not f_dom or not f_emp:
        log("[PULADO] Arquivos não encontrados.")
        return

    df_d = preparar_dataframe(ler_arquivo_na_rede(f_dom), "DOMINIO")
    df_e = preparar_dataframe(ler_arquivo_na_rede(f_emp), "EMPRESA")

    if df_d.empty and df_e.empty: 
        log("[ERRO] Dados insuficientes.")
        return

    # Cruzamento
    df_final = pd.merge(df_d, df_e, on="Nota", how="outer", suffixes=('_Dom', '_Emp'), indicator=True)
    
    # Preenche vazios
    df_final['Valor_Dom'] = df_final['Valor_Dom'].fillna(0.0)
    df_final['Valor_Emp'] = df_final['Valor_Emp'].fillna(0.0)
    
    # Unifica código (prioridade Domínio)
    if 'Codigo_Dom' in df_final.columns:
        df_final['Codigo'] = df_final['Codigo_Dom'].fillna(df_final['Codigo_Emp'])
    elif 'Codigo' not in df_final.columns:
        df_final['Codigo'] = ""
    
    df_final["Diferenca"] = df_final["Valor_Dom"] - df_final["Valor_Emp"]
    
    df_final["Status"] = df_final.apply(lambda r: 
        "Só Domínio" if r['_merge'] == 'left_only' else 
        ("Só Empresa" if r['_merge'] == 'right_only' else 
        ("Divergência Valor" if abs(r['Diferenca']) > 0.05 else "OK")), axis=1)
    
    # Seleciona Colunas Finais
    cols_finais = ['Codigo', 'Nota', 'Valor_Dom', 'Valor_Emp', 'Diferenca', 'Status']
    df_final = df_final[[c for c in cols_finais if c in df_final.columns]]
    
    try:
        df_final["k"] = pd.to_numeric(df_final["Nota"])
        df_final.sort_values("k", inplace=True)
        df_final.drop(columns="k", inplace=True)
    except: df_final.sort_values("Nota", inplace=True)

    # Salva
    os.makedirs(DIR_SAIDA_RPA, exist_ok=True)
    fname = f"Conciliacao_{empresa.replace(' ', '_')}_{mes_ano}.xlsx"
    fout = os.path.join(DIR_SAIDA_RPA, fname)
    
    while True:
        try:
            with pd.ExcelWriter(fout, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Resultado')
                wb, ws = writer.book, writer.sheets['Resultado']
                fmt_m = wb.add_format({'num_format': '#,##0.00'})
                fmt_r = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                fmt_y = wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})
                fmt_b = wb.add_format({'bg_color': '#BDD7EE', 'font_color': '#000000'})
                
                ws.set_column('A:B', 12)
                ws.set_column('C:E', 18, fmt_m)
                ws.set_column('F:F', 25)
                
                ws.conditional_format('F2:F9999', {'type':'text', 'criteria':'containing', 'value':'Divergência', 'format':fmt_r})
                ws.conditional_format('F2:F9999', {'type':'text', 'criteria':'containing', 'value':'Só Domínio', 'format':fmt_y})
                ws.conditional_format('F2:F9999', {'type':'text', 'criteria':'containing', 'value':'Só Empresa', 'format':fmt_b})

            log(f"    [SUCESSO] Salvo: {fname}")
            break
        except PermissionError:
            log(f"\n    !!! ARQUIVO ABERTO: {fname} !!!")
            input("    Feche o Excel e pressione ENTER...")
        except Exception as e:
            log(f"    [ERRO SALVAR] {e}")
            break

def extrair_ano(mes_ano):
    try:
        return mes_ano.split("-")[1]
    except Exception:
        return ""

def resolver_bases(mes_ano):
    ano = extrair_ano(mes_ano)
    bases = []
    for tmpl in BASES_TEMPLATE:
        if "{mes_ano}" in tmpl or "{ano}" in tmpl:
            bases.append(tmpl.format(ano=ano, mes_ano=mes_ano))
        else:
            bases.append(tmpl)
    return bases

if __name__ == "__main__":
    mes_ano_cli = sys.argv[1] if len(sys.argv) > 1 else MES_ANO_DEFAULT
    empresas_cli = sys.argv[2:] if len(sys.argv) > 2 else []
    alvo_empresas = empresas_cli if empresas_cli else LISTA_EMPRESAS

    log(f"--- INICIANDO RPA (V34 - CORRETOR DE COLUNA) [{mes_ano_cli}] ---")
    log(f"Empresas alvo: {alvo_empresas}")
    base = None
    for p in resolver_bases(mes_ano_cli):
        if os.path.exists(p): base = p; break
    if base:
        for e in alvo_empresas: processar_empresa(e, base, mes_ano_cli)
    else:
        log("[ERRO FATAL] Pasta base n??o encontrada.")
    log("\n--- FIM ---")
