import os
import re
import sys
import shutil
import tempfile
import subprocess
import configparser
from utils import resource_path
from pathlib import Path
from numbers import Integral
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
    Detecta cabecalho e recorta colunas relevantes.
    Dom: Nota col 4, Valor col 20, Data col 2
    Emp: Nota col 12, Valor col 17, Data col 10, Status Nfe (quando existir)
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame(columns=["Nota", "Valor", "Data", "Codigo", "Status_NFE"])

    def _find_header_row(df: pd.DataFrame, must_have: List[str], max_rows: int = 30) -> Optional[int]:
        lim = min(max_rows, len(df))
        for i in range(lim):
            row = df.iloc[i].astype(str).str.lower()
            if all(row.str.contains(t, na=False, regex=False).any() for t in must_have):
                return i
        return None

    if isinstance(df_raw.columns[0], Integral):
        # Alguns relatórios do Domínio vêm com cabeçalho em linha fixa (5),
        # e alguns relatórios de Empresa trazem cabeçalho por volta da linha 3.
        if len(df_raw) <= 6:
            log("[ERRO] Planilha sem linhas suficientes para cabecalho")
            return pd.DataFrame(columns=["Codigo", "Nota", "Valor", "Data", "Status_NFE"])

        if tipo_origem == "EMPRESA":
            header_idx = _find_header_row(df_raw, must_have=["n.nota", "status nfe"], max_rows=20)
            if header_idx is None:
                header_idx = _find_header_row(df_raw, must_have=["n.nota", "status"], max_rows=25)
            if header_idx is None:
                header_idx = _find_header_row(df_raw, must_have=["nota", "status"], max_rows=25)
            if header_idx is None:
                header_idx = 5
        else:
            header_idx = _find_header_row(df_raw, must_have=["nota", "valor"], max_rows=20)
            if header_idx is None:
                header_idx = 5

        df_raw.columns = df_raw.iloc[header_idx].astype(str).str.lower().str.strip()
        df_raw = df_raw.iloc[header_idx + 1 :].reset_index(drop=True)
    else:
        df_raw.columns = df_raw.columns.astype(str).str.lower().str.strip()

    mask_total = df_raw.apply(lambda r: r.astype(str).str.contains("total", case=False, na=False)).any(axis=1)
    if mask_total.any():
        df_raw = df_raw.loc[~mask_total].reset_index(drop=True)

    if tipo_origem == "DOMINIO":
        if len(df_raw.columns) <= 22:
            log("[ERRO] DOMINIO: colunas insuficientes")
            return pd.DataFrame(columns=["Codigo", "Nota", "Valor", "Data", "Status_NFE"])
        col_nota = next(
            (c for c in df_raw.columns if isinstance(c, str) and c.strip() == "nota"),
            None,
        ) or df_raw.columns[4]
        col_data = next(
            (c for c in df_raw.columns if isinstance(c, str) and c.strip() == "data"),
            None,
        ) or df_raw.columns[2]
        col_valor = next(
            (c for c in df_raw.columns if isinstance(c, str) and "valor cont" in c),
            None,
        ) or df_raw.columns[20]
        col_cod = None
        col_status = None
        df_raw = cortar_inicio(df_raw, df_raw.columns.get_loc(col_nota))
        valor_series = df_raw[col_valor]
        data_series = parse_data(df_raw[col_data])
    else:
        if len(df_raw.columns) <= 17:
            log("[ERRO] EMPRESA: colunas insuficientes")
            return pd.DataFrame(columns=["Codigo", "Nota", "Valor", "Data", "Status_NFE"])
        col_nota = next(
            (c for c in df_raw.columns if isinstance(c, str) and "n.nota" in c),
            None,
        ) or df_raw.columns[12]
        # Preferência: "Total Nota" (valor total do documento). Se não existir, tenta "Total Produtos".
        col_valor = (
            next((c for c in df_raw.columns if isinstance(c, str) and "total nota" in c), None)
            or next((c for c in df_raw.columns if isinstance(c, str) and "total produtos" in c), None)
            or (df_raw.columns[17] if len(df_raw.columns) > 17 else df_raw.columns[-1])
        )
        col_data = next(
            (c for c in df_raw.columns if isinstance(c, str) and ("dt.emiss" in c or "dt.emissão" in c)),
            None,
        ) or df_raw.columns[10]
        col_cod = None
        df_raw = cortar_inicio(df_raw, df_raw.columns.get_loc(col_nota))
        valor_series = df_raw[col_valor]
        data_series = parse_data(df_raw[col_data])
        # Status NFe fica na coluna U (indice 20) no relatório Empresa.
        col_status = df_raw.columns[20] if len(df_raw.columns) > 20 else None
        if col_status is None:
            col_status = next(
                (c for c in df_raw.columns if isinstance(c, str) and "status" in c and "nfe" in c),
                None,
            )

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
        if col_status:
            df_new["Status_NFE"] = df_raw[col_status]
        else:
            df_new["Status_NFE"] = ""
    except Exception as exc:
        log(f"[ERRO] Recorte de colunas: {exc}")
        return pd.DataFrame(columns=["Codigo", "Nota", "Valor", "Data", "Status_NFE"])

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

    def agregar_por_nota(df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame(columns=["Codigo", "Nota", "Valor", "Data", "Status_NFE"])
        out = df.copy()
        for c in ["Codigo", "Nota", "Valor", "Data", "Status_NFE"]:
            if c not in out.columns:
                out[c] = ""

        def first_non_empty(series: pd.Series):
            for v in series:
                if pd.notna(v) and str(v).strip() != "":
                    return v
            return ""

        grouped = (
            out.groupby("Nota", as_index=False)
            .agg(
                Valor=("Valor", "sum"),
                Data=("Data", "min"),
                Codigo=("Codigo", first_non_empty),
                Status_NFE=("Status_NFE", first_non_empty),
            )
        )
        return grouped[["Codigo", "Nota", "Valor", "Data", "Status_NFE"]]

    # Saida agora na pasta da empresa: .../RELATORIO RPA - <empresa>/Conciliacao
    out_dir = path_rpa / "Conciliacao"
    os.makedirs(out_dir, exist_ok=True)
    fout = out_dir / f"Conciliacao_{empresa.replace(' ', '_')}_{mes_ano}.xlsx"

    try:
        with pd.ExcelWriter(fout, engine="xlsxwriter") as writer:
            wb = writer.book
            fmt_header = wb.add_format(
                {
                    "bold": True,
                    "bg_color": "#D9E1F2",
                    "border": 1,
                    "text_wrap": True,
                    "align": "center",
                    "valign": "vcenter",
                }
            )
            fmt_text = wb.add_format({"text_wrap": True, "align": "center", "valign": "vcenter"})
            fmt_date = wb.add_format({"num_format": "dd/mm/yyyy", "align": "center", "valign": "vcenter"})
            fmt_m = wb.add_format({"num_format": "#,##0.00", "align": "center", "valign": "vcenter"})
            fmt_r = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "align": "center", "valign": "vcenter"})
            fmt_y = wb.add_format({"bg_color": "#FFEB9C", "font_color": "#9C6500", "align": "center", "valign": "vcenter"})
            fmt_b = wb.add_format({"bg_color": "#BDD7EE", "font_color": "#000000", "align": "center", "valign": "vcenter"})

            # Ordem das abas: Resumo -> Conciliacao Completa -> Inutilizadas
            ws_r = wb.add_worksheet("Resumo")
            ws = wb.add_worksheet("Conciliacao Completa")
            writer.sheets["Resumo"] = ws_r
            writer.sheets["Conciliacao Completa"] = ws
            # A mesma Nota pode aparecer múltiplas vezes (ex.: por CFOP). Conciliação é feita por Nota,
            # somando os valores para obter o total por documento.
            df_d_g = agregar_por_nota(df_d) if not df_d.empty else pd.DataFrame(columns=["Codigo", "Nota", "Valor", "Data", "Status_NFE"])
            df_e_g_full = agregar_por_nota(df_e) if not df_e.empty else pd.DataFrame(columns=["Codigo", "Nota", "Valor", "Data", "Status_NFE"])
            log(f"Notas únicas (Dom/Emp): {len(df_d_g)} / {len(df_e_g_full)}")

            # Se a empresa tem Status NFE, separa notas inutilizadas (ex.: "I") em aba dedicada.
            df_inutilizadas = pd.DataFrame()
            df_e_g_all = df_e_g_full.copy()
            df_e_g = df_e_g_full.copy()
            if "Status_NFE" in df_e_g.columns and not df_e_g.empty:
                status_norm = df_e_g["Status_NFE"].astype(str).str.strip().str.upper()
                mask_inut = status_norm.eq("I") | status_norm.str.startswith("I ")
                if mask_inut.any():
                    df_inutilizadas = df_e_g.loc[mask_inut].copy()
                    df_e_g = df_e_g.loc[~mask_inut].copy()

                    notas_inut = set(df_inutilizadas["Nota"].astype(str))
                    if "Nota" in df_d_g.columns and not df_d_g.empty:
                        df_d_g = df_d_g.loc[~df_d_g["Nota"].astype(str).isin(notas_inut)].copy()
                    log(f"Notas inutilizadas (empresa): {len(df_inutilizadas)}")

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

            # Reinsere inutilizadas no Resultado com status próprio (para não aparecer como "So Empresa")
            if not df_inutilizadas.empty:
                n_inut = len(df_inutilizadas)
                cod_inut = df_inutilizadas["Codigo"] if "Codigo" in df_inutilizadas.columns else pd.Series([""] * n_inut)
                nota_inut = df_inutilizadas["Nota"] if "Nota" in df_inutilizadas.columns else pd.Series([""] * n_inut)
                val_inut = df_inutilizadas["Valor"] if "Valor" in df_inutilizadas.columns else pd.Series([0.0] * n_inut)
                df_inut_res = pd.DataFrame(
                    {
                        "Codigo": cod_inut,
                        "Nota": nota_inut,
                        "Valor_Dom": 0.0,
                        "Valor_Emp": val_inut,
                        "Diferenca": 0.0 - val_inut,
                        "Status": "Inutilizada",
                    }
                )
                df_final = pd.concat([df_final, df_inut_res], ignore_index=True)

            try:
                df_final["k"] = pd.to_numeric(df_final["Nota"])
                df_final.sort_values("k", inplace=True)
                df_final.drop(columns="k", inplace=True)
            except Exception:
                df_final.sort_values("Nota", inplace=True)

            # Aba de resumo para leitura rápida
            total_resultado = len(df_final)
            qtd_inutilizadas = int((df_final["Status"] == "Inutilizada").sum()) if "Status" in df_final.columns else 0
            qtd_so_empresa = int((df_final["Status"] == "So Empresa").sum()) if "Status" in df_final.columns else 0
            qtd_so_dominio = int((df_final["Status"] == "So Dominio").sum()) if "Status" in df_final.columns else 0
            qtd_ok = int((df_final["Status"] == "OK").sum()) if "Status" in df_final.columns else 0
            qtd_div = int((df_final["Status"] == "Divergencia Valor").sum()) if "Status" in df_final.columns else 0

            df_resumo = pd.DataFrame(
                [
                    ["Empresa", empresa],
                    ["Mes/Ano", mes_ano],
                    ["Notas (Conciliacao Completa)", total_resultado],
                    ["Inutilizadas", qtd_inutilizadas],
                    ["So Empresa", qtd_so_empresa],
                    ["So Dominio", qtd_so_dominio],
                    ["OK", qtd_ok],
                    ["Divergencia Valor", qtd_div],
                    ["Notas lidas (Dom/Emp)", f"{len(df_d_g)} / {len(df_e_g_all)}"],
                ],
                columns=["Item", "Valor"],
            )
            df_resumo.to_excel(writer, index=False, sheet_name="Resumo")
            ws_r.set_row(0, 22)
            ws_r.write(0, 0, "Item", fmt_header)
            ws_r.write(0, 1, "Valor", fmt_header)
            ws_r.freeze_panes(1, 0)
            ws_r.autofilter(0, 0, max(0, len(df_resumo)), 1)
            ws_r.set_column("A:A", 28, fmt_text)
            ws_r.set_column("B:B", 40, fmt_text)

            # Aba principal (conciliacao completa)
            df_final.to_excel(writer, index=False, sheet_name="Conciliacao Completa")
            ws.set_row(0, 22)
            for col_idx, col_name in enumerate(df_final.columns.tolist()):
                ws.write(0, col_idx, col_name, fmt_header)
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(0, len(df_final)), max(0, len(df_final.columns) - 1))

            ws.set_column("A:A", 14, fmt_text)
            ws.set_column("B:B", 12, fmt_text)
            ws.set_column("C:E", 18, fmt_m)
            ws.set_column("F:F", 22, fmt_text)
            ws.conditional_format("F2:F9999", {"type": "text", "criteria": "containing", "value": "Divergencia", "format": fmt_r})
            ws.conditional_format("F2:F9999", {"type": "text", "criteria": "containing", "value": "So Dominio", "format": fmt_y})
            ws.conditional_format("F2:F9999", {"type": "text", "criteria": "containing", "value": "So Empresa", "format": fmt_b})
            ws.conditional_format("F2:F9999", {"type": "text", "criteria": "containing", "value": "Inutilizada", "format": fmt_b})

            if not df_inutilizadas.empty:
                ws2 = wb.add_worksheet("Inutilizadas")
                writer.sheets["Inutilizadas"] = ws2
                cols_inut = ["Codigo", "Nota", "Data", "Valor", "Status_NFE"]
                df_inut_out = df_inutilizadas[[c for c in cols_inut if c in df_inutilizadas.columns]].copy()
                try:
                    df_inut_out["k"] = pd.to_numeric(df_inut_out["Nota"], errors="coerce")
                    df_inut_out.sort_values("k", inplace=True)
                    df_inut_out.drop(columns="k", inplace=True)
                except Exception:
                    pass
                df_inut_out.to_excel(writer, index=False, sheet_name="Inutilizadas")
                ws2.set_row(0, 22)
                for col_idx, col_name in enumerate(df_inut_out.columns.tolist()):
                    ws2.write(0, col_idx, col_name, fmt_header)
                ws2.freeze_panes(1, 0)
                ws2.autofilter(0, 0, max(0, len(df_inut_out)), max(0, len(df_inut_out.columns) - 1))
                ws2.set_column("A:A", 14, fmt_text)
                ws2.set_column("B:B", 12, fmt_text)
                ws2.set_column("C:C", 14, fmt_date)
                ws2.set_column("D:D", 18, fmt_m)
                ws2.set_column("E:E", 12, fmt_text)
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
