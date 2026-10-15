#!/usr/bin/env python3
"""
Automação de notas fiscais → planilha de contas a pagar.

Como funciona:
  1. Você (ou o e-mail, ou o sistema do fornecedor) salva o XML da NF-e na pasta "entrada".
  2. O script lê a nota e lança uma linha por parcela na planilha "contas_a_pagar.xlsx",
     já com fornecedor, CNPJ, número da nota, valor e vencimento.
  3. A nota vai para a pasta "processadas". Se algo der errado, vai para "erros"
     e o motivo fica registrado no arquivo "processamento.log".

Uso:
  python automacao_nf.py            -> processa o que estiver na pasta e termina
  python automacao_nf.py --vigiar   -> fica rodando e processa cada nota que chegar

Requisito: Python 3.9+ e a biblioteca openpyxl (pip install openpyxl).
"""

import argparse
import datetime as dt
import logging
import shutil
import sys
import time
import xml.etree.ElementTree as ET
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------- configuração
BASE = Path(__file__).resolve().parent
PASTA_ENTRADA = BASE / "entrada"
PASTA_PROCESSADAS = BASE / "processadas"
PASTA_ERROS = BASE / "erros"
PLANILHA = BASE / "contas_a_pagar.xlsx"
ARQUIVO_LOG = BASE / "processamento.log"
INTERVALO_VIGIA_SEG = 5
LINHAS_FORMATADAS = 5000  # alcance das regras de cor e da lista "Pago?"

NS = {"nfe": "http://www.portalfiscal.inf.br/nfe"}
ABA = "Contas a pagar"
COLUNAS = [  # (título, largura)
    ("Lançado em", 12), ("Fornecedor", 34), ("CNPJ", 20), ("NF", 9), ("Emissão", 11),
    ("Parcela", 9), ("Vencimento", 12), ("Valor (R$)", 14), ("Situação", 16), ("Pago?", 8),
    ("Chave de acesso", 48), ("Arquivo", 30),
]
# letras das colunas usadas em fórmulas
C_VENC, C_VALOR, C_SIT, C_PAGO, C_CHAVE, C_PARC = "G", "H", "I", "J", "K", "F"

FONTE = "Arial"
AZUL = "1F4E79"

log = logging.getLogger("nf")


# ---------------------------------------------------------------- leitura da NF-e
class NotaInvalida(Exception):
    pass


def _txt(no, caminho):
    if no is None:
        return None
    achado = no.find(caminho, NS)
    return achado.text.strip() if achado is not None and achado.text else None


def _data(valor):
    """Aceita '2026-10-03', '2026-10-03T10:15:00-03:00'."""
    if not valor:
        return None
    return dt.date.fromisoformat(valor[:10])


def _cnpj_formatado(c):
    if c and len(c) == 14:
        return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"
    if c and len(c) == 11:
        return f"{c[:3]}.{c[3:6]}.{c[6:9]}-{c[9:]}"
    return c or ""


def ler_nfe(caminho: Path) -> dict:
    try:
        raiz = ET.parse(caminho).getroot()
    except ET.ParseError as e:
        raise NotaInvalida(f"XML ilegível ({e})")

    inf = raiz.find(".//nfe:infNFe", NS)
    if inf is None:
        raise NotaInvalida("não é uma NF-e (tag infNFe não encontrada)")

    chave = _txt(raiz, ".//nfe:protNFe/nfe:infProt/nfe:chNFe") or (inf.get("Id") or "").replace("NFe", "")
    ide, emit, total = inf.find("nfe:ide", NS), inf.find("nfe:emit", NS), inf.find("nfe:total/nfe:ICMSTot", NS)

    nota = {
        "chave": chave,
        "numero": _txt(ide, "nfe:nNF"),
        "emissao": _data(_txt(ide, "nfe:dhEmi") or _txt(ide, "nfe:dEmi")),
        "fornecedor": _txt(emit, "nfe:xNome"),
        "cnpj": _cnpj_formatado(_txt(emit, "nfe:CNPJ") or _txt(emit, "nfe:CPF")),
        "valor_total": float(_txt(total, "nfe:vNF") or 0),
        "parcelas": [],
    }
    if not nota["fornecedor"] or not nota["numero"]:
        raise NotaInvalida("nota sem fornecedor ou sem número")

    for dup in inf.findall("nfe:cobr/nfe:dup", NS):
        nota["parcelas"].append({
            "numero": _txt(dup, "nfe:nDup") or "",
            "vencimento": _data(_txt(dup, "nfe:dVenc")),
            "valor": float(_txt(dup, "nfe:vDup") or 0),
        })

    # Nota sem duplicatas: lança o valor total sem vencimento, marcado para conferir.
    if not nota["parcelas"]:
        nota["parcelas"].append({"numero": "única", "vencimento": None, "valor": nota["valor_total"]})
    return nota


# ---------------------------------------------------------------- planilha
def _borda():
    fino = Side(style="thin", color="D9D9D9")
    return Border(bottom=fino)


def criar_planilha(caminho: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = ABA
    for i, (titulo, largura) in enumerate(COLUNAS, start=1):
        c = ws.cell(row=1, column=i, value=titulo)
        c.font = Font(name=FONTE, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[c.column_letter].width = largura
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=len(COLUNAS)).column_letter}1"

    fim = LINHAS_FORMATADAS
    faixa = f"{C_SIT}2:{C_SIT}{fim}"
    cores = [("Vencida", "F8CBAD", "9C0006"), ("Vence em 7 dias", "FFE699", "7F6000"),
             ("Paga", "C6EFCE", "006100"), ("Verificar", "D9D9D9", "404040")]
    for texto, fundo, letra in cores:
        ws.conditional_formatting.add(faixa, FormulaRule(
            formula=[f'${C_SIT}2="{texto}"'], fill=PatternFill("solid", fgColor=fundo), font=Font(color=letra, bold=True)))

    dv = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=True)
    dv.prompt, dv.promptTitle = "Marque Sim quando a parcela for paga.", "Pago?"
    ws.add_data_validation(dv)
    dv.add(f"{C_PAGO}2:{C_PAGO}{fim}")
    ws[f"{C_PAGO}1"].fill = PatternFill("solid", fgColor="BF8F00")  # coluna para o usuário preencher

    # Aba de resumo: só fórmulas, atualiza sozinha.
    r = wb.create_sheet("Resumo")
    r.column_dimensions["A"].width = 30
    r.column_dimensions["B"].width = 18
    r.column_dimensions["C"].width = 12
    r["A1"] = "Contas a pagar — resumo"
    r["A1"].font = Font(name=FONTE, bold=True, size=14, color=AZUL)
    r["A2"] = "Atualiza sozinho a partir da aba 'Contas a pagar'. Situação calculada com a data de hoje."
    r["A2"].font = Font(name=FONTE, italic=True, size=9, color="808080")
    for col, t in zip("ABC", ["Situação", "Valor (R$)", "Parcelas"]):
        c = r[f"{col}4"]
        c.value = t
        c.font = Font(name=FONTE, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=AZUL)
    aba = f"'{ABA}'"
    linhas = ["Vencida", "Vence em 7 dias", "A vencer", "Verificar", "Paga"]
    for i, s in enumerate(linhas, start=5):
        r[f"A{i}"] = s
        r[f"B{i}"] = f'=SUMIFS({aba}!${C_VALOR}:${C_VALOR},{aba}!${C_SIT}:${C_SIT},A{i})'
        r[f"C{i}"] = f'=COUNTIFS({aba}!${C_SIT}:${C_SIT},A{i})'
    t = 5 + len(linhas)
    r[f"A{t}"] = "Em aberto (sem as pagas)"
    r[f"B{t}"] = f"=SUM(B5:B{t - 2})"
    r[f"C{t}"] = f"=SUM(C5:C{t - 2})"
    for col in "ABC":
        r[f"{col}{t}"].font = Font(name=FONTE, bold=True)
        r[f"{col}{t}"].border = Border(top=Side(style="thin", color="000000"))
    for row in r.iter_rows(min_row=5, max_row=t):
        for c in row:
            if c.font.name != FONTE or not c.font.bold:
                c.font = Font(name=FONTE, bold=c.font.bold)
            if c.column_letter == "B":
                c.number_format = '#,##0.00;-#,##0.00;"-"'
    wb.save(caminho)


def lancar(ws, nota: dict, arquivo: str, vistos: set) -> int:
    hoje = dt.date.today()
    novas = 0
    for p in nota["parcelas"]:
        id_parcela = (nota["chave"], str(p["numero"]))
        if id_parcela in vistos:
            continue
        linha = ws.max_row + 1
        valores = [hoje, nota["fornecedor"], nota["cnpj"], nota["numero"], nota["emissao"],
                   p["numero"], p["vencimento"], p["valor"], None, None, nota["chave"], arquivo]
        for col, v in enumerate(valores, start=1):
            c = ws.cell(row=linha, column=col, value=v)
            c.font = Font(name=FONTE)
            c.border = _borda()
        for col in ("A", "E", C_VENC):
            ws[f"{col}{linha}"].number_format = "DD/MM/YYYY"
        ws[f"{C_VALOR}{linha}"].number_format = "#,##0.00"
        ws[f"{C_SIT}{linha}"] = (
            f'=IF({C_PAGO}{linha}="Sim","Paga",IF({C_VENC}{linha}="","Verificar",'
            f'IF({C_VENC}{linha}<TODAY(),"Vencida",IF({C_VENC}{linha}-TODAY()<=7,"Vence em 7 dias","A vencer"))))'
        )
        ws[f"{C_SIT}{linha}"].font = Font(name=FONTE)
        ws[f"{C_SIT}{linha}"].alignment = Alignment(horizontal="center")
        ws[f"{C_PAGO}{linha}"].alignment = Alignment(horizontal="center")
        vistos.add(id_parcela)
        novas += 1
    return novas


# ---------------------------------------------------------------- execução
def mover(arquivo: Path, destino: Path):
    destino.mkdir(exist_ok=True)
    alvo = destino / arquivo.name
    if alvo.exists():
        alvo = destino / f"{arquivo.stem}_{dt.datetime.now():%Y%m%d%H%M%S}{arquivo.suffix}"
    shutil.move(str(arquivo), str(alvo))


def processar_pasta() -> int:
    for p in (PASTA_ENTRADA, PASTA_PROCESSADAS, PASTA_ERROS):
        p.mkdir(exist_ok=True)
    arquivos = sorted(f for f in PASTA_ENTRADA.iterdir() if f.is_file() and not f.name.startswith("."))
    if not arquivos:
        return 0
    if not PLANILHA.exists():
        criar_planilha(PLANILHA)
        log.info("Planilha criada: %s", PLANILHA.name)

    try:
        wb = load_workbook(PLANILHA)
    except PermissionError:
        log.warning("A planilha está aberta em outro programa. Feche-a; tento de novo na próxima rodada.")
        return 0
    ws = wb[ABA]
    vistos = {(str(k), str(p)) for p, *_meio, k in
              ws.iter_rows(min_row=2, min_col=6, max_col=11, values_only=True) if k}

    total = 0
    for arq in arquivos:
        if arq.suffix.lower() != ".xml":
            log.error("%s: não é XML. Peça o XML ao fornecedor (vem junto com o PDF da nota).", arq.name)
            mover(arq, PASTA_ERROS)
            continue
        try:
            nota = ler_nfe(arq)
        except NotaInvalida as e:
            log.error("%s: %s", arq.name, e)
            mover(arq, PASTA_ERROS)
            continue
        novas = lancar(ws, nota, arq.name, vistos)
        if novas:
            log.info("%s: NF %s de %s — %d parcela(s), R$ %s", arq.name, nota["numero"], nota["fornecedor"],
                     novas, f"{nota['valor_total']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        else:
            log.info("%s: NF %s já estava na planilha, ignorada.", arq.name, nota["numero"])
        mover(arq, PASTA_PROCESSADAS)
        total += novas

    ws.auto_filter.ref = f"A1:L{max(ws.max_row, 1)}"
    try:
        wb.save(PLANILHA)
    except PermissionError:
        log.warning("Não consegui salvar: a planilha está aberta. Feche-a e rode de novo.")
    return total


def main():
    # Garante acentos corretos no terminal do Windows.
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    ap = argparse.ArgumentParser(description="Lança NF-e (XML) na planilha de contas a pagar.")
    ap.add_argument("--vigiar", action="store_true", help="fica rodando e processa cada nota que chegar")
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s", datefmt="%d/%m %H:%M:%S",
                        handlers=[logging.FileHandler(ARQUIVO_LOG, encoding="utf-8"), logging.StreamHandler(sys.stdout)])
    if args.vigiar:
        log.info("Vigiando a pasta '%s' (Ctrl+C para parar)…", PASTA_ENTRADA.name)
        try:
            while True:
                processar_pasta()
                time.sleep(INTERVALO_VIGIA_SEG)
        except KeyboardInterrupt:
            log.info("Encerrado.")
    else:
        n = processar_pasta()
        log.info("Pronto: %d parcela(s) lançada(s).", n)


if __name__ == "__main__":
    main()
