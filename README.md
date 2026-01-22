# Documentacao - RPA - Dominio x Empresa

## 1. Visao Geral
Automacao para conciliar relatorios de notas fiscais entre Dominio e Empresa. A UI permite escolher empresa e mes/ano e gera um Excel consolidado com o status por nota.

---

## 2. Estrutura do Projeto
```
RPA - Dominio x Empresa/
|-- main.py
|-- front_base.py
|-- conciliacao.py
|-- config.ini
|-- utils.py
```

---

## 3. Componentes Principais
- main.py: ponto de entrada da aplicacao.
- front_base.py: UI Tkinter (empresa + mes/ano + progresso).
- conciliacao.py: leitura dos arquivos, conciliacao e exportacao do Excel.

---

## 4. Fluxo de Trabalho
1) Usuario escolhe empresa e mes/ano.
2) Sistema localiza arquivos DOMINIO/EMPRESA.
3) Converte .xls para .xlsx se necessario.
4) Gera Conciliacao_<empresa>_<mes_ano>.xlsx.

---
 
## 5. Configuracao
config.ini:
- [GERAL]: PASTA_BASE, ARQUIVOS_GLOBAIS, MES_ANO.
- [EMPRESAS]: caminhos por empresa.
- [PADROES]: nomes dos arquivos Dominio/Empresa.
- [estrutura_relatorios]: subpasta dos relatorios.

---

## 6. Empresas Suportadas
As empresas sao definidas em [EMPRESAS] do config.ini.

---

## 7. Formato dos Arquivos de Entrada
Relatorios em .xls/.xlsx. A logica usa indices fixos:
- DOMINIO: Nota col 4, Valor col 20, Data col 2.
- EMPRESA: Nota col 12, Valor col 17, Data col 10.
- Cabecalho esperado na linha 6.

---

## 8. Relatorio de Saida
Arquivo Excel na subpasta Conciliacao, com colunas:
Codigo, Nota, Valor_Dom, Valor_Emp, Diferenca, Status.
Status inclui: OK, So Dominio, So Empresa, Divergencia Valor.

---

## 9. Dependencias
- Python 3
- pandas, openpyxl, xlsxwriter
- LibreOffice (soffice/scalc) para conversao de .xls

---

## 10. Execucao
```bash
python main.py
# ou
python conciliacao.py 11-2025 "DROGARIA LIMEIRA"
```

---

## 11. Mecanismos de Seguranca
- Pula empresas sem pasta ou sem arquivos.
- Tolerancia de diferenca ajustavel no config.ini.

---

## 12. Troubleshooting
- LibreOffice nao encontrado: instale e ajuste o caminho.
- Arquivos nao localizados: revise keywords e nomes no config.ini.

---

## 13. Historico de Versoes
Nao mantido.

---

## 14. Notas de Seguranca
Caminhos de rede precisam estar acessiveis.
