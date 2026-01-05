import os
import re
import sys
import shutil
import tempfile
import subprocess
import configparser
from utils import resource_path
from pathlib import Path
from typing import Callable, Optional, List, Dict

import pandas as pd

# --- Config carregada do config.ini ---
CFG_PATH = Path(resource_path("config.ini"))
CFG = configparser.ConfigParser()
CFG.optionxform = str  # preserva maiusculas/minusculas das chaves
if CFG_PATH.exists():
    CFG.read(CFG_PATH, encoding="utf-8")

# Config padrao (sobrescrito pelo ini quando presente)
MES_ANO_DEFAULT = CFG.get("GERAL", "MES_ANO", fallback="11-2025")
DIR_SAIDA_RPA = CFG.get("GERAL", "ARQUIVOS_GLOBAIS", fallback=CFG.get("GERAL", "PASTA_BASE", fallback=r"V:\Fiscal\RPA"))

# Bases de busca para as empresas (suporta uso de {ano} e {mes_ano})
BASES_TEMPLATE = [v for _, v in CFG.items("caminhos_base")] if CFG.has_section("caminhos_base") else []

# Estrutura de relatorios
SUBPASTA_RELATORIO = CFG.get(
    "estrutura_relatorios", "subpasta_relatorio", fallback=r""
)

# Palavras-chave para identificar arquivos
def _keyword_from_cfg(cfg_value: str, default: str) -> str:
    if not cfg_value:
        return default
    u = cfg_value.upper()
    if "DOMINIO" in u:
        return "DOMINIO"
    if "EMPRESA" in u and default == "EMPRESA":
        return "EMPRESA"
    return default


KEYWORD_DOMINIO = _keyword_from_cfg(CFG.get("estrutura_relatorios", "arquivo_dominio", fallback=""), "DOMINIO")
KEYWORD_EMPRESA = _keyword_from_cfg(CFG.get("estrutura_relatorios", "arquivo_empresa", fallback=""), "EMPRESA")

# --- Helpers de config (novo .ini) ---
def _expand_vars(value: str, empresa: str = "", mes_ano: str = "") -> str:
    if not value:
        return value
    ctx: Dict[str, str] = {}
    if CFG.has_section("GERAL"):
        for k, v in CFG["GERAL"].items():
            ctx[k.lower()] = v
    ctx["empresa"] = empresa
    ctx["mes_ano"] = mes_ano
    ctx["ano"] = extrair_ano(mes_ano)
    try:
        return value.format(**{k.upper(): v for k, v in ctx.items()}, **ctx)
    except Exception:
        try:
            return value.format(**ctx)
        except Exception:
            return value


def carregar_empresas_cfg(mes_ano: str) -> Dict[str, Dict[str, str]]:
    """
    Retorna dict: nome -> {base_dir, arquivo_dom, arquivo_emp}
    Usa seções [EMPRESAS], [PADROES] e [PADROES.<NOME>].
    """
    if not CFG.has_section("EMPRESAS"):
        return {}

    base_padrao_dom = CFG.get("PADROES", "ARQUIVO_DOMINIO", fallback="")
    base_padrao_emp = CFG.get("PADROES", "ARQUIVO_EMPRESA", fallback="")

    empresas_cfg: Dict[str, Dict[str, str]] = {}
    for nome, caminho in CFG["EMPRESAS"].items():
        nome_limpo = nome.strip()
        base_dir = _expand_vars(caminho, empresa=nome_limpo, mes_ano=mes_ano)

        sec_esp = f"PADROES.{nome_limpo}"
        arq_dom = CFG.get(sec_esp, "ARQUIVO_DOMINIO", fallback=base_padrao_dom)
        arq_emp = CFG.get(sec_esp, "ARQUIVO_EMPRESA", fallback=base_padrao_emp)

        empresas_cfg[nome_limpo] = {
            "base_dir": base_dir,
            "arquivo_dom": _expand_vars(arq_dom, empresa=nome_limpo, mes_ano=mes_ano),
            "arquivo_emp": _expand_vars(arq_emp, empresa=nome_limpo, mes_ano=mes_ano),
        }
    return empresas_cfg


LIBREOFFICE_CANDIDATOS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files\LibreOffice\program\scalc.exe",
]

LOG_FN: Optional[Callable[[str], None]] = None


def set_logger(fn: Callable[[str], None]):
    """Define callback para registrar mensagens (UI)."""
    global LOG_FN
    LOG_FN = fn


def log(msg: str):
    if LOG_FN:
        try:
            LOG_FN(msg)
        except Exception:
            pass
    print(msg)


def converter_para_float(texto):
    if pd.isna(texto) or str(texto).strip() == "":
        return 0.0
    if isinstance(texto, (int, float)):
        return float(texto)
    t = str(texto).strip()
    t = re.sub(r"[^\d.,-]", "", t)
    if not t:
        return 0.0
    if "," in t and "." in t:
        if t.find(",") > t.find("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        val = float(t)
        return val
    except Exception:
        return 0.0


def normalizar_nota(nota):
    if pd.isna(nota):
        return "S/N"
    try:
        so_numeros = re.sub(r"[^\d.]", "", str(nota))
        if not so_numeros:
            return "S/N"
        val = float(so_numeros)
        if val <= 0:
            return "S/N"
        return str(int(val))
    except Exception:
        return str(nota).strip()


def preencher_mesclados(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    try:
        return df.ffill()
    except Exception:
        return df


def parse_data(col) -> pd.Series:
    s = col.copy()
    mask_header = s.astype(str).str.contains("dt", case=False, na=False) | s.astype(str).str.contains("emiss", case=False, na=False)
    s = s.mask(mask_header)
    ser = pd.to_datetime(s, errors="coerce", format="%d/%m/%Y")
    if ser.isna().mean() > 0.5:
        ser = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if ser.isna().mean() > 0.5:
        ser = pd.to_datetime(s, errors="coerce", unit="d", origin="1899-12-30")
    return ser


def encontrar_libreoffice() -> Optional[Path]:
    for c in LIBREOFFICE_CANDIDATOS:
        p = Path(c)
        if p.exists():
            return p
    return None


def converter_para_xlsx(caminho_arquivo: Path) -> Optional[Path]:
    if caminho_arquivo.suffix.lower() != ".xls":
        return caminho_arquivo

    libre = encontrar_libreoffice()
    if not libre:
        log("[ERRO] LibreOffice nao encontrado (soffice/scalc).")
        return None

    xlsx_dir = caminho_arquivo.parent / "XLSX"
    xlsx_dir.mkdir(exist_ok=True)
    destino = xlsx_dir / f"{caminho_arquivo.stem}.xlsx"

    tmpdir = Path(tempfile.mkdtemp(prefix="conv_rpa_"))
    cmd = [
        str(libre),
        "--headless",
        "--convert-to",
        "xlsx",
        "--outdir",
        str(tmpdir),
        str(caminho_arquivo),
    ]
    log(f"Convertendo {caminho_arquivo.name} para XLSX...")
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    except Exception as exc:
        log(f"[ERRO CONVERSAO] {exc}")
        shutil.rmtree(tmpdir, ignore_errors=True)
        return None

    if res.returncode != 0:
        log(f"[ERRO CONVERSAO] {res.stderr.strip() or res.stdout.strip()}")
        shutil.rmtree(tmpdir, ignore_errors=True)
        return None

    candidatos = sorted(tmpdir.glob(f"{caminho_arquivo.stem}*.xlsx"), key=os.path.getmtime, reverse=True)
    if not candidatos:
        log("[ERRO CONVERSAO] Nenhum .xlsx gerado.")
        shutil.rmtree(tmpdir, ignore_errors=True)
        return None

    try:
        shutil.move(str(candidatos[0]), destino)
    except Exception as exc:
        log(f"[ERRO CONVERSAO] Falha ao mover arquivo convertido: {exc}")
        shutil.rmtree(tmpdir, ignore_errors=True)
        return None

    shutil.rmtree(tmpdir, ignore_errors=True)
    return destino


def ler_arquivo(caminho_arquivo: Path) -> Optional[pd.DataFrame]:
    log(f"Lendo arquivo: {caminho_arquivo.name}")
    caminho_para_ler = converter_para_xlsx(caminho_arquivo)
    if not caminho_para_ler:
        log("[ERRO] Conversao/obtencao do arquivo falhou.")
        return None
    try:
        df = pd.read_excel(caminho_para_ler, header=None, engine="openpyxl")
        return preencher_mesclados(df)
    except Exception as exc:
        log(f"[ERRO LEITURA] {exc}")
        return None


def cortar_inicio(df: pd.DataFrame, col_nota_idx: int) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    start_idx = 0
    for i in range(len(df)):
        nota_raw = df.iat[i, col_nota_idx] if col_nota_idx < df.shape[1] else None
        nota_norm = normalizar_nota(nota_raw)
        if nota_norm not in ("S/N", "0"):
            start_idx = i
            break
    if start_idx > 0:
        return df.iloc[start_idx:].reset_index(drop=True)
    return df


def preparar_dataframe(df_raw: pd.DataFrame, tipo_origem: str) -> pd.DataFrame:
    """
    Titulos na linha 6 (indice 5).
    Dom: Nota col 4, Valor col 20, Data col 2
    Emp: Nota col 12, Valor col 17, Data col 10
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame(columns=["Nota", "Valor"])

    if isinstance(df_raw.columns[0], int):
        if len(df_raw) <= 6:
            log("[ERRO] Planilha sem linhas suficientes para cabecalho")
            return pd.DataFrame(columns=["Codigo", "Nota", "Valor"])
        df_raw.columns = df_raw.iloc[5].astype(str).str.lower().str.strip()
        df_raw = df_raw.iloc[6:].reset_index(drop=True)
    else:
        df_raw.columns = df_raw.columns.astype(str).str.lower().str.strip()

    mask_total = df_raw.apply(lambda r: r.astype(str).str.contains("total", case=False, na=False)).any(axis=1)
    if mask_total.any():
        df_raw = df_raw.loc[~mask_total].reset_index(drop=True)

    if tipo_origem == "DOMINIO":
        if len(df_raw.columns) <= 22:
            log("[ERRO] DOMINIO: colunas insuficientes")
            return pd.DataFrame(columns=["Codigo", "Nota", "Valor"])
        col_nota = df_raw.columns[4]
        col_valor = df_raw.columns[20]
        col_data = df_raw.columns[2]
        col_cod = None
        df_raw = cortar_inicio(df_raw, df_raw.columns.get_loc(col_nota))
        valor_series = df_raw[col_valor]
        data_series = parse_data(df_raw[col_data])
    else:
        if len(df_raw.columns) <= 17:
            log("[ERRO] EMPRESA: colunas insuficientes")
            return pd.DataFrame(columns=["Codigo", "Nota", "Valor"])
        col_nota = df_raw.columns[12]
        col_valor = df_raw.columns[17]
        col_data = df_raw.columns[10]
        col_cod = None
        df_raw = cortar_inicio(df_raw, df_raw.columns.get_loc(col_nota))
        valor_series = df_raw[col_valor]
        data_series = parse_data(df_raw[col_data])

    try:
        df_new = pd.DataFrame({
            "Nota": df_raw[col_nota],
            "Valor": valor_series,
            "Data": data_series,
        })
        if col_cod:
            df_new["Codigo"] = df_raw[col_cod]
        else:
            df_new["Codigo"] = ""
    except Exception as exc:
        log(f"[ERRO] Recorte de colunas: {exc}")
        return pd.DataFrame(columns=["Codigo", "Nota", "Valor"])

    df_new["Nota"] = df_new["Nota"].apply(normalizar_nota)
    df_new["Valor"] = df_new["Valor"].apply(converter_para_float)

    df_new["Nota_num"] = pd.to_numeric(df_new["Nota"], errors="coerce")
    df_new = df_new[
        df_new["Nota_num"].notna()
        & (df_new["Valor"] > 0.01)
    ]

    return df_new.drop(columns=["Nota_num"])


def extrair_ano(mes_ano: str) -> str:
    try:
        return mes_ano.split("-")[1]
    except Exception:
        return ""


def resolver_bases(mes_ano: str) -> List[str]:
    ano = extrair_ano(mes_ano)
    if not BASES_TEMPLATE:
        return []
    resolved = []
    for tmpl in BASES_TEMPLATE:
        try:
            resolved.append(tmpl.format(ano=ano, mes_ano=mes_ano))
        except Exception:
            resolved.append(tmpl)
    return resolved


def processar_empresa(empresa: str, pasta_base: str, mes_ano: str, arquivo_dom: Optional[str] = None, arquivo_emp: Optional[str] = None):
    log(f"Empresa: {empresa}")
    # Calcula caminho da pasta que contem os relatorios para a empresa.
    # Se SUBPASTA_RELATORIO tiver placeholder {empresa}, usa diretamente.
    # Caso contrario, adiciona "RELATORIO RPA - {empresa}" ao final.
    try:
        sub_rel_str = SUBPASTA_RELATORIO.format(empresa=empresa)
    except Exception:
        sub_rel_str = SUBPASTA_RELATORIO

    if not sub_rel_str:
        sub_rel_path = Path(".")
    else:
        sub_rel_path = Path(sub_rel_str)
        if "{empresa}" not in SUBPASTA_RELATORIO and "RELATORIO RPA" not in sub_rel_path.name.upper():
            sub_rel_path = sub_rel_path / f"RELATORIO RPA - {empresa}"

    base_path = Path(pasta_base)
    if (base_path / empresa).exists():
        base_path = base_path / empresa
    path_rpa = base_path / sub_rel_path

    if not path_rpa.exists():
        log("[PULADO] Pasta nao encontrada.")
        return

    # Garante pasta XLSX para conversoes.
    xlsx_dir = path_rpa / "XLSX"
    xlsx_dir.mkdir(exist_ok=True)

    dom_candidates = []
    emp_candidates = []

    def tentar_adicionar_por_nome(arq_nome: Optional[str], destino: list):
        if not arq_nome:
            return
        alvo = path_rpa / arq_nome
        if alvo.exists():
            destino.append(alvo)
        else:
            # tenta procurar por nome exato dentro da pasta (qualquer subpasta direta)
            for f in path_rpa.glob("**/*"):
                if f.is_file() and f.name.lower() == arq_nome.lower():
                    destino.append(f)
                    break

    tentar_adicionar_por_nome(arquivo_dom, dom_candidates)
    tentar_adicionar_por_nome(arquivo_emp, emp_candidates)

    # Se nao achar pelos nomes especificos, recorre ao padrao por keyword
    for f in path_rpa.iterdir():
        if f.name.startswith("~$"):
            continue
        up = f.name.upper()
        if KEYWORD_DOMINIO in up and f.suffix.lower() in (".xls", ".xlsx"):
            dom_candidates.append(f)
        if KEYWORD_EMPRESA in up and f.suffix.lower() in (".xls", ".xlsx"):
            emp_candidates.append(f)
    xlsx_dir = path_rpa / "XLSX"
    if xlsx_dir.exists():
        for f in xlsx_dir.iterdir():
            if f.name.startswith("~$"):
                continue
            up = f.name.upper()
            if KEYWORD_DOMINIO in up and f.suffix.lower() == ".xlsx":
                dom_candidates.append(f)
            if KEYWORD_EMPRESA in up and f.suffix.lower() == ".xlsx":
                emp_candidates.append(f)

    def escolher_arquivos(files):
        escolhidos = {}
        for f in files:
            stem = f.stem.lower()
            if stem in escolhidos:
                if f.suffix.lower() == ".xlsx":
                    escolhidos[stem] = f
            else:
                escolhidos[stem] = f
        return list(escolhidos.values())

    dom_files = escolher_arquivos(dom_candidates)
    emp_files = escolher_arquivos(emp_candidates)

    if not dom_files or not emp_files:
        log("[PULADO] Arquivos DOMINIO/EMPRESA nao encontrados.")
        return

    dom_files = sorted(dom_files)
    emp_files = sorted(emp_files)

    dfs_dom = []
    for f in dom_files:
        log(f"Lendo DOMINIO: {f.name}")
        dfs_dom.append(preparar_dataframe(ler_arquivo(f), "DOMINIO"))
    dfs_emp = []
    for f in emp_files:
        log(f"Lendo EMPRESA: {f.name}")
        dfs_emp.append(preparar_dataframe(ler_arquivo(f), "EMPRESA"))

    df_d = pd.concat([d for d in dfs_dom if d is not None], ignore_index=True) if dfs_dom else pd.DataFrame()
    df_e = pd.concat([e for e in dfs_emp if e is not None], ignore_index=True) if dfs_emp else pd.DataFrame()

    if df_d.empty and df_e.empty:
        log("[ERRO] Dados insuficientes.")
        return

    # Saida agora na pasta da empresa: .../RELATORIO RPA - <empresa>/Conciliacao
    out_dir = path_rpa / "Conciliacao"
    os.makedirs(out_dir, exist_ok=True)
    fout = out_dir / f"Conciliacao_{empresa.replace(' ', '_')}_{mes_ano}.xlsx"

    try:
        with pd.ExcelWriter(fout, engine="xlsxwriter") as writer:
            df_d_g = df_d.drop_duplicates(subset="Nota", keep="first") if not df_d.empty else pd.DataFrame(columns=["Nota", "Valor", "Codigo"])
            df_e_g = df_e.drop_duplicates(subset="Nota", keep="first") if not df_e.empty else pd.DataFrame(columns=["Nota", "Valor", "Codigo"])
            log(f"Notas lidas Dom/Emp: {len(df_d_g)} / {len(df_e_g)}")

            df_final = pd.merge(df_d_g, df_e_g, on="Nota", how="outer", suffixes=("_Dom", "_Emp"), indicator=True)
            df_final["Valor_Dom"] = df_final["Valor_Dom"].fillna(0.0)
            df_final["Valor_Emp"] = df_final["Valor_Emp"].fillna(0.0)
            df_final["Codigo"] = df_final.get("Codigo_Dom", pd.Series()).fillna(df_final.get("Codigo_Emp", ""))
            df_final["Diferenca"] = df_final["Valor_Dom"] - df_final["Valor_Emp"]
            df_final["Status"] = df_final.apply(
                lambda r: "So Dominio"
                if r["_merge"] == "left_only"
                else ("So Empresa" if r["_merge"] == "right_only" else ("Divergencia Valor" if abs(r["Diferenca"]) > 0.05 else "OK")),
                axis=1,
            )

            cols_finais = ["Codigo", "Nota", "Valor_Dom", "Valor_Emp", "Diferenca", "Status"]
            df_final = df_final[[c for c in cols_finais if c in df_final.columns]]

            try:
                df_final["k"] = pd.to_numeric(df_final["Nota"])
                df_final.sort_values("k", inplace=True)
                df_final.drop(columns="k", inplace=True)
            except Exception:
                df_final.sort_values("Nota", inplace=True)

            df_final.to_excel(writer, index=False, sheet_name="Resultado")
            wb, ws = writer.book, writer.sheets["Resultado"]
            fmt_m = wb.add_format({"num_format": "#,##0.00"})
            fmt_r = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
            fmt_y = wb.add_format({"bg_color": "#FFEB9C", "font_color": "#9C6500"})
            fmt_b = wb.add_format({"bg_color": "#BDD7EE", "font_color": "#000000"})
            ws.set_column("A:B", 12)
            ws.set_column("C:E", 18, fmt_m)
            ws.set_column("F:F", 25)
            ws.conditional_format("F2:F9999", {"type": "text", "criteria": "containing", "value": "Divergencia", "format": fmt_r})
            ws.conditional_format("F2:F9999", {"type": "text", "criteria": "containing", "value": "So Dominio", "format": fmt_y})
            ws.conditional_format("F2:F9999", {"type": "text", "criteria": "containing", "value": "So Empresa", "format": fmt_b})
        log(f"Consolidado salvo: {fout}")
    except Exception as exc:
        log(f"[ERRO SALVAR] {exc}")


def run_conciliacao(mes_ano: str, empresas: List[str]):
    log(f"Iniciando conciliacao [{mes_ano}]")

    empresas_cfg = carregar_empresas_cfg(mes_ano)

    # Se ini define empresas com caminhos especificos, usa eles.
    if empresas_cfg:
        alvo = empresas or list(empresas_cfg.keys())
        for emp in alvo:
            conf = empresas_cfg.get(emp)
            if not conf:
                log(f"[PULADO] Empresa nao configurada no ini: {emp}")
                continue
            base_dir = conf.get("base_dir") or ""
            if not base_dir:
                log(f"[PULADO] Base nao informada para {emp}")
                continue
            log(f"Base: {base_dir}")
            processar_empresa(
                emp,
                base_dir,
                mes_ano,
                arquivo_dom=conf.get("arquivo_dom"),
                arquivo_emp=conf.get("arquivo_emp"),
            )
        log("Fim")
        return

    # Fallback antigo (usa caminhos_base + subpastas)
    base = None
    for p in resolver_bases(mes_ano):
        if os.path.exists(p):
            base = p
            break
    if not base:
        log("[ERRO FATAL] Pasta base nao encontrada.")
        return
    log(f"Base: {base}")
    for emp in empresas:
        processar_empresa(emp, base, mes_ano)
    log("Fim")


if __name__ == "__main__":
    mes_ano_cli = sys.argv[1] if len(sys.argv) > 1 else MES_ANO_DEFAULT
    empresas_cli = sys.argv[2:] if len(sys.argv) > 2 else []
    if not empresas_cli and CFG.has_section("empresas"):
        empresas_cli = list(CFG["empresas"].values())
    if not empresas_cli:
        empresas_cli = ["DROGARIA LIMEIRA", "DROGARIA MORELLI FILIAL", "DROGARIA MORELLI MTZ"]
    run_conciliacao(mes_ano_cli, empresas_cli)
