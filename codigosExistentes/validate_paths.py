"""
Validador basico de caminhos para o RPA de conciliacao.
Uso:
    python validate_paths.py 11-2025 "DROGARIA LIMEIRA" "DROGARIA LIMEIRA FILIAL"
Se empresas nao forem informadas, usa a lista padrao.
"""

import os
import sys

MES_ANO_DEFAULT = "11-2025"

LISTA_EMPRESAS_PADRAO = [
    "DROGARIA LIMEIRA",
    "DROGARIA MORELLI FILIAL",
    "DROGARIA MORELLI MTZ",
]

BASES_TEMPLATE = [
    r"N:\Matriz-Jds\Arquivos NF-e\{ano}\{mes_ano}",
    r"N:\Arquivos NF-e\{ano}\{mes_ano}",
    r"N:\{ano}\{mes_ano}",
]

RELATORIO_SUB = "ESCRITA FISCAL"


def extrair_ano(mes_ano: str) -> str:
    try:
        return mes_ano.split("-")[1]
    except Exception:
        return ""


def resolver_bases(mes_ano: str):
    ano = extrair_ano(mes_ano)
    bases = []
    for tmpl in BASES_TEMPLATE:
        if "{mes_ano}" in tmpl or "{ano}" in tmpl:
            bases.append(tmpl.format(ano=ano, mes_ano=mes_ano))
        else:
            bases.append(tmpl)
    return bases


def validar_caminho_empresa(base, empresa):
    pasta_relatorio = f"RELATORIO RPA - {empresa}"
    path_rpa = os.path.join(base, empresa, RELATORIO_SUB, pasta_relatorio)
    existe = os.path.exists(path_rpa)
    arquivos = []
    if existe:
        try:
            arquivos = os.listdir(path_rpa)
        except Exception as exc:
            return existe, [], f"Erro ao listar: {exc}"
    return existe, arquivos, ""


def main():
    mes_ano = sys.argv[1] if len(sys.argv) > 1 else MES_ANO_DEFAULT
    empresas = sys.argv[2:] if len(sys.argv) > 2 else LISTA_EMPRESAS_PADRAO

    print(f"Mes/Ano: {mes_ano}")
    print(f"Empresas: {empresas}")

    base_encontrada = None
    for b in resolver_bases(mes_ano):
        print(f"Testando base: {b}")
        if os.path.exists(b):
            base_encontrada = b
            print("  -> encontrada")
            break
        else:
            print("  -> nao existe")

    if not base_encontrada:
        print("Nenhuma base encontrada. Ajuste BASES_TEMPLATE ou o mes/ano.")
        sys.exit(1)

    print(f"\nBase usada: {base_encontrada}\n")
    ok = True
    for emp in empresas:
        existe, arquivos, erro = validar_caminho_empresa(base_encontrada, emp)
        print(f"[{emp}]")
        if not existe:
            print("  Caminho nao encontrado.")
            ok = False
            continue
        if erro:
            print(f"  {erro}")
            ok = False
            continue
        dom = [f for f in arquivos if "DOMINIO" in f.upper() and f.lower().endswith((".xls", ".xlsx"))]
        emp_arq = [f for f in arquivos if "EMPRESA" in f.upper() and f.lower().endswith((".xls", ".xlsx"))]
        print(f"  Pasta OK: {len(arquivos)} arquivo(s)")
        print(f"  DOMINIO: {dom or 'nao localizado'}")
        print(f"  EMPRESA: {emp_arq or 'nao localizado'}")
        if not dom or not emp_arq:
            ok = False

    if ok:
        print("\nStatus: OK - todos os caminhos/arquivos minimos encontrados.")
    else:
        print("\nStatus: pendencias encontradas (veja itens acima).")


if __name__ == "__main__":
    main()
