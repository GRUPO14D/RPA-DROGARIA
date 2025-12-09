import os
import re
import sys
import shutil
import tempfile
import subprocess
from pathlib import Path
from typing import Callable, Optional, List

import pandas as pd

# Config padrao
MES_ANO_DEFAULT = "11-2025"
TETO_VALOR_ACEITAVEL = 5_000_000.00

BASES_TEMPLATE = [
    r"N:\Matriz-Jds\Arquivos NF-e\{ano}\{mes_ano}",
    r"N:\Arquivos NF-e\{ano}\{mes_ano}",
    r"N:\{ano}\{mes_ano}",
]

DIR_SAIDA_RPA = r"V:\Fiscal\RPA"
KEYWORD_DOMINIO = "DOMINIO"
KEYWORD_EMPRESA = "EMPRESA"

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
    return [tmpl.format(ano=ano, mes_ano=mes_ano) for tmpl in BASES_TEMPLATE]


def processar_empresa(empresa: str, pasta_base: str, mes_ano: str):
    log(f"Empresa: {empresa}")
    pasta_relatorio = f"RELATORIO RPA - {empresa}"
    path_rpa = Path(pasta_base) / empresa / "ESCRITA FISCAL" / pasta_relatorio

    if not path_rpa.exists():
        log("[PULADO] Pasta nao encontrada.")
        return

    dom_candidates = []
    emp_candidates = []
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

    os.makedirs(DIR_SAIDA_RPA, exist_ok=True)
    fout = Path(DIR_SAIDA_RPA) / f"Conciliacao_{empresa.replace(' ', '_')}_{mes_ano}.xlsx"

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
    if not empresas_cli:
        empresas_cli = ["DROGARIA LIMEIRA", "DROGARIA MORELLI FILIAL", "DROGARIA MORELLI MTZ"]
    run_conciliacao(mes_ano_cli, empresas_cli)
