"""
Robo que busca o levantamento semanal de precos de combustiveis da ANP,
pega a planilha mais recente disponibilizada no site oficial, e extrai a
media nacional de Gasolina, Etanol e Diesel para o arquivo anp-referencia.json.

Se algo falhar (ANP fora do ar, mudanca de layout, etc), o script NAO apaga
o arquivo antigo - assim o app sempre tem um numero de referencia, mesmo
que nao seja o mais novo daquela semana.
"""
import json
import re
import sys
from datetime import datetime, timezone

import requests

PAGINA_ANP = "https://www.gov.br/anp/pt-br/assuntos/precos-e-defesa-da-concorrencia/precos/levantamento-de-precos-de-combustiveis-ultimas-semanas-pesquisadas"
SAIDA = "anp-referencia.json"

# nomes de produto como costumam aparecer nas planilhas da ANP
ALIAS_PRODUTOS = {
    "gasolina": ["GASOLINA COMUM", "GASOLINA C COMUM", "GASOLINA"],
    "etanol": ["ETANOL HIDRATADO", "ETANOL"],
    "diesel": ["OLEO DIESEL S10", "ÓLEO DIESEL S10", "OLEO DIESEL", "ÓLEO DIESEL"],
}


def achar_link_planilha_mais_recente():
    resp = requests.get(PAGINA_ANP, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    matches = re.findall(r'href="([^"]*resumo[_-]semanal[_-]lpc[^"]*\.xlsx)"', resp.text, re.IGNORECASE)
    if not matches:
        raise RuntimeError("Nao encontrei nenhum link de planilha na pagina da ANP. O layout pode ter mudado.")
    link = matches[0]
    if link.startswith("http"):
        return link
    return "https://www.gov.br" + link


def baixar_planilha(url):
    resp = requests.get(url, timeout=60, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    nome_arquivo = "planilha_anp.xlsx"
    with open(nome_arquivo, "wb") as f:
        f.write(resp.content)
    return nome_arquivo


def extrair_medias_nacionais(caminho_planilha):
    import openpyxl

    wb = openpyxl.load_workbook(caminho_planilha, data_only=True)
    resultado = {}
    resultado_fallback = {}

    for aba in wb.sheetnames:
        ws = wb[aba]
        linhas = list(ws.iter_rows(values_only=True))

        for linha in linhas:
            if not linha:
                continue
            textos = [str(c).strip().upper() if isinstance(c, str) else "" for c in linha]
            tem_brasil = "BRASIL" in textos

            for idx, texto in enumerate(textos):
                if not texto:
                    continue
                chave_produto = None
                for chave, aliases in ALIAS_PRODUTOS.items():
                    if any(texto == a or texto.startswith(a) for a in aliases):
                        chave_produto = chave
                        break
                if chave_produto is None:
                    continue

                preco_achado = None
                for c in linha[idx + 1:]:
                    if isinstance(c, (int, float)) and 0.5 <= float(c) <= 20:
                        preco_achado = round(float(c), 3)
                        break
                if preco_achado is None:
                    continue

                if tem_brasil and chave_produto not in resultado:
                    resultado[chave_produto] = preco_achado
                elif chave_produto not in resultado_fallback:
                    resultado_fallback[chave_produto] = preco_achado

        if len(resultado) == 3:
            break

    for chave in ALIAS_PRODUTOS:
        if chave not in resultado and chave in resultado_fallback:
            resultado[chave] = resultado_fallback[chave]

    faltando = [k for k in ALIAS_PRODUTOS if k not in resultado]
    if faltando:
        raise RuntimeError(f"Nao consegui achar preco nacional para: {faltando}. Pode ser preciso ajustar o script pro layout atual da planilha.")

    return resultado


def carregar_json_existente():
    try:
        with open(SAIDA, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return None


def main():
    try:
        link = achar_link_planilha_mais_recente()
        print(f"Planilha encontrada: {link}")
        caminho = baixar_planilha(link)
        medias = extrair_medias_nacionais(caminho)

        dados = {
            "gasolina": medias["gasolina"],
            "etanol": medias["etanol"],
            "diesel": medias["diesel"],
            "fonte": link,
            "atualizado_em": datetime.now(timezone.utc).isoformat(),
        }
        with open(SAIDA, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
        print("anp-referencia.json atualizado com sucesso:", dados)

    except Exception as e:
        print(f"ERRO ao atualizar dados da ANP: {e}", file=sys.stderr)
        anterior = carregar_json_existente()
        if anterior:
            print("Mantendo o arquivo anterior sem alteracoes.")
        else:
            dados = {
                "gasolina": 6.09, "etanol": 4.42, "diesel": 6.31,
                "fonte": "valor inicial (robo ainda nao conseguiu buscar da ANP)",
                "atualizado_em": datetime.now(timezone.utc).isoformat(),
            }
            with open(SAIDA, "w", encoding="utf-8") as f:
                json.dump(dados, f, ensure_ascii=False, indent=2)
        sys.exit(1)


if __name__ == "__main__":
    main()
