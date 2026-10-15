#!/usr/bin/env python3
"""Gera notas fiscais FICTÍCIAS (XML no layout NF-e 4.00) na pasta 'entrada', para demonstrar a automação.
Empresas, CNPJs e chaves são inventados — não valem como documento fiscal."""

import datetime as dt
import random
from pathlib import Path

PASTA = Path(__file__).resolve().parent / "entrada"
random.seed(2026)


def dv_cnpj(base12):
    def dig(nums, pesos):
        s = sum(int(n) * p for n, p in zip(nums, pesos)) % 11
        return "0" if s < 2 else str(11 - s)
    d1 = dig(base12, [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])
    d2 = dig(base12 + d1, [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])
    return base12 + d1 + d2


def cnpj_ficticio():
    return dv_cnpj("".join(random.choice("0123456789") for _ in range(8)) + "0001")


def chave(cnpj, emissao, numero):
    base = f"35{emissao:%y%m}{cnpj}55001{numero:09d}1{random.randint(0, 99999999):08d}"
    pesos, s = [2, 3, 4, 5, 6, 7, 8, 9], 0
    for i, n in enumerate(reversed(base)):
        s += int(n) * pesos[i % 8]
    dv = 11 - s % 11
    return base + str(0 if dv >= 10 else dv)


def xml_nfe(fornecedor, numero, emissao, itens, parcelas_dias):
    cnpj = cnpj_ficticio()
    ch = chave(cnpj, emissao, numero)
    total = round(sum(q * v for _, q, v in itens), 2)
    dets = "".join(
        f'<det nItem="{i}"><prod><cProd>{i:04d}</cProd><xProd>{d}</xProd><qCom>{q}</qCom>'
        f'<vUnCom>{v:.2f}</vUnCom><vProd>{q * v:.2f}</vProd></prod></det>'
        for i, (d, q, v) in enumerate(itens, start=1))
    cobr = ""
    if parcelas_dias:
        n = len(parcelas_dias)
        valores = [round(total / n, 2)] * n
        valores[-1] = round(total - sum(valores[:-1]), 2)
        dups = "".join(
            f"<dup><nDup>{i:03d}</nDup><dVenc>{emissao + dt.timedelta(days=d):%Y-%m-%d}</dVenc><vDup>{v:.2f}</vDup></dup>"
            for i, (d, v) in enumerate(zip(parcelas_dias, valores), start=1))
        cobr = f"<cobr><fat><nFat>{numero}</nFat><vOrig>{total:.2f}</vOrig><vLiq>{total:.2f}</vLiq></fat>{dups}</cobr>"
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00"><NFe><infNFe versao="4.00" Id="NFe' + ch + '">'
        f"<ide><cUF>35</cUF><natOp>Venda de mercadoria</natOp><mod>55</mod><serie>1</serie><nNF>{numero}</nNF>"
        f"<dhEmi>{emissao:%Y-%m-%d}T09:30:00-03:00</dhEmi></ide>"
        f"<emit><CNPJ>{cnpj}</CNPJ><xNome>{fornecedor}</xNome></emit>"
        "<dest><CNPJ>11222333000181</CNPJ><xNome>VERTICE DISTRIBUIDORA LTDA (FICTICIA)</xNome></dest>"
        f"{dets}<total><ICMSTot><vProd>{total:.2f}</vProd><vNF>{total:.2f}</vNF></ICMSTot></total>{cobr}"
        f"</infNFe></NFe><protNFe><infProt><chNFe>{ch}</chNFe><cStat>100</cStat></infProt></protNFe></nfeProc>"
    )


NOTAS = [  # fornecedor, nº, dias atrás da emissão, itens (descrição, qtd, preço), prazos das parcelas em dias
    ("Torrefação Serra Alta Ltda", 48211, 40, [("Café torrado 500g", 1200, 11.40)], [30, 60]),
    ("Embalagens Rio Claro S.A.", 9174, 25, [("Caixa de papelão 40x30", 3000, 2.15), ("Fita adesiva 48mm", 400, 4.90)], [21]),
    ("Laticínios Vale do Sol Ltda", 33058, 12, [("Queijo muçarela kg", 380, 31.50)], [14, 28, 42]),
    ("Transportes Rota Sul Ltda", 1502, 9, [("Frete rodoviário — setembro", 1, 8740.00)], [10]),
    ("Grãos do Cerrado Comércio Ltda", 77420, 5, [("Arroz tipo 1 5kg", 900, 19.80), ("Feijão carioca 1kg", 1500, 6.10)], [28, 56, 84]),
    ("Papelaria Central ME", 2231, 3, [("Material de escritório", 1, 486.30)], []),  # sem duplicata → "Verificar"
]


def main():
    PASTA.mkdir(exist_ok=True)
    hoje = dt.date.today()
    for forn, num, atras, itens, parc in NOTAS:
        xml = xml_nfe(forn, num, hoje - dt.timedelta(days=atras), itens, parc)
        (PASTA / f"NFe_{num}.xml").write_text(xml, encoding="utf-8")
    # cópia repetida da primeira nota (o script não lança duas vezes) e um PDF (vai para "erros")
    (PASTA / "NFe_48211_reenviada.xml").write_text((PASTA / "NFe_48211.xml").read_text(encoding="utf-8"), encoding="utf-8")
    (PASTA / "DANFE_scan_fornecedor.pdf").write_bytes(b"%PDF-1.4\n% arquivo de exemplo\n")
    print(f"{len(NOTAS) + 2} arquivos de exemplo criados em '{PASTA.name}'.")


if __name__ == "__main__":
    main()
