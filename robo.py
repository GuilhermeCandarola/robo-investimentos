#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ROBO DE INVESTIMENTOS
Busca dados de mercado, le a sua carteira, gera o Painel em Excel e envia alertas.
Feito para rodar sozinho a cada 15 minutos pelo Agendador de Tarefas do Windows.

Uso:
    python robo.py            -> execucao normal (dados reais)
    python robo.py --demo     -> dados ficticios, sem internet (apenas para teste)
    python robo.py --teste-alerta -> envia uma mensagem de teste pelos canais ativos

IMPORTANTE: os rankings e sinais sao calculos objetivos sobre dados publicos.
Nao constituem recomendacao de investimento.
"""
import configparser
import datetime as dt
import hashlib
import io
import json
import logging
import math
import os
import re
import smtplib
import sys
import urllib.parse
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor
from email.message import EmailMessage
from email.utils import parsedate_to_datetime
from pathlib import Path

import numpy as np
import pandas as pd
import requests
from lxml import html as lxml_html
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
CACHE = BASE / "cache"
DEMO = "--demo" in sys.argv
AGORA = dt.datetime.now()
HOJE = AGORA.date()
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")

log = logging.getLogger("robo")
STATUS = []  # (fonte, situacao, detalhe)


# ----------------------------------------------------------------- utilidades
def configurar_log():
    arq = BASE / "robo.log"
    if arq.exists() and arq.stat().st_size > 1_000_000:
        arq.unlink()
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(arq, encoding="utf-8"), logging.StreamHandler()])


def carregar_config():
    cfg = configparser.ConfigParser(interpolation=None)
    cfg.read(BASE / "config.ini", encoding="utf-8-sig")
    return cfg


def cfg_f(cfg, sec, chave, padrao):
    try:
        return float(cfg.get(sec, chave).replace(",", "."))
    except Exception:
        return padrao


def cfg_sim(cfg, sec, chave):
    return cfg.get(sec, chave, fallback="nao").strip().lower() in ("sim", "s", "true", "1", "yes")


def br_num(txt):
    """'1.234,56' -> 1234.56 ; '10,84%' -> 0.1084 ; vazio -> NaN"""
    if txt is None:
        return math.nan
    s = str(txt).strip()
    if s in ("", "-"):
        return math.nan
    pct = s.endswith("%")
    s = s.replace("%", "").replace(".", "").replace(",", ".")
    try:
        v = float(s)
    except ValueError:
        return math.nan
    return v / 100 if pct else v


def fbr(v, dec=2, sinal=False):
    """Numero no padrao brasileiro: 1.234,56"""
    s = f"{v:+,.{dec}f}" if sinal else f"{v:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fpc(v, dec=1, sinal=False):
    return fbr(v * 100, dec, sinal) + "%"


def http_get(url, timeout=25, tentativas=2):
    if DEMO:
        import demo_dados
        return demo_dados.resposta(url)
    erro = None
    for _ in range(tentativas):
        try:
            r = requests.get(url, headers={"User-Agent": UA, "Accept-Language": "pt-BR,pt;q=0.9"},
                             timeout=timeout)
            r.raise_for_status()
            return r.content
        except Exception as e:  # noqa
            erro = e
    raise erro


def com_cache(nome, func):
    """Executa func(); se falhar, usa o ultimo resultado salvo em cache."""
    CACHE.mkdir(exist_ok=True)
    arq = CACHE / f"{nome}.pkl"
    try:
        dado = func()
        if dado is None or (hasattr(dado, "empty") and dado.empty):
            raise ValueError("retorno vazio")
        pd.to_pickle(dado, arq)
        STATUS.append((nome, "OK", f"atualizado {AGORA:%d/%m %H:%M}"))
        return dado
    except Exception as e:  # noqa
        log.warning("Falha em %s: %s", nome, e)
        if arq.exists():
            quando = dt.datetime.fromtimestamp(arq.stat().st_mtime)
            STATUS.append((nome, "CACHE", f"falhou agora; usando dados de {quando:%d/%m %H:%M}"))
            return pd.read_pickle(arq)
        STATUS.append((nome, "FALHA", str(e)[:120]))
        return None


# ------------------------------------------------------------ fontes de dados
def _linhas_tabela(conteudo):
    doc = lxml_html.fromstring(conteudo.decode("latin-1", errors="replace"))
    tabelas = doc.xpath("//table")
    if not tabelas:
        raise ValueError("tabela nao encontrada")
    tab = max(tabelas, key=lambda t: len(t.xpath(".//tr")))
    linhas = []
    for tr in tab.xpath(".//tr"):
        tds = tr.xpath("./td")
        if tds:
            linhas.append([td.text_content().strip() for td in tds])
    return linhas


COLS_ACOES = ["Ticker", "Preco", "PL", "PVP", "PSR", "DY", "P_Ativo", "P_CapGiro", "P_EBIT",
              "P_AtivCircLiq", "EV_EBIT", "EV_EBITDA", "MrgBruta", "MrgEbit", "MrgLiq",
              "LiqCorr", "ROIC", "ROE", "Liquidez", "PatrimLiq", "DivLiq_Patrim", "Cresc5a"]
COLS_FIIS = ["Ticker", "Segmento", "Preco", "FFOYield", "DY", "PVP", "ValorMercado", "Liquidez",
             "QtdImoveis", "PrecoM2", "AluguelM2", "CapRate", "Vacancia"]


def buscar_acoes():
    linhas = _linhas_tabela(http_get("https://www.fundamentus.com.br/resultado.php"))
    dados = [l[:22] for l in linhas if len(l) >= 22]
    df = pd.DataFrame(dados, columns=COLS_ACOES)
    for c in COLS_ACOES[1:]:
        df[c] = df[c].map(br_num)
    return df[df["Preco"] > 0].reset_index(drop=True)


def buscar_fiis():
    linhas = _linhas_tabela(http_get("https://www.fundamentus.com.br/fii_resultado.php"))
    dados = [l[:13] for l in linhas if len(l) >= 13]
    df = pd.DataFrame(dados, columns=COLS_FIIS)
    for c in COLS_FIIS[2:]:
        df[c] = df[c].map(br_num)
    return df[df["Preco"] > 0].reset_index(drop=True)


def json_get(url, timeout=25, tentativas=3):
    """JSON com novas tentativas (o BCB as vezes devolve pagina de erro com status 200)."""
    import time
    erro = None
    for i in range(tentativas):
        try:
            return json.loads(http_get(url, timeout=timeout, tentativas=1))
        except Exception as e:  # noqa
            erro = e
            if not DEMO and i < tentativas - 1:
                time.sleep(2 * (i + 1))
    raise erro


def sgs_ultimos(codigo, n=1):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{codigo}/dados/ultimos/{n}?formato=json"
    return json_get(url)


def sgs_periodo(codigo, d0, d1):
    vazio = pd.DataFrame({"data": [], "valor": []})
    if d0 >= d1:  # aplicacao de hoje: ainda nao ha nenhum dia fechado
        return vazio
    url = (f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{codigo}/dados?formato=json"
           f"&dataInicial={d0:%d/%m/%Y}&dataFinal={d1:%d/%m/%Y}")
    try:
        js = json_get(url, timeout=40)
    except requests.HTTPError as e:
        if e.response is not None and e.response.status_code == 404:
            return vazio  # o BCB responde 404 quando o periodo nao tem nenhum dado
        raise
    if not js:
        return vazio
    df = pd.DataFrame(js)
    df["data"] = pd.to_datetime(df["data"], format="%d/%m/%Y").dt.date
    df["valor"] = df["valor"].astype(float)
    return df


def buscar_macro():
    series = {"Selic meta (% a.a.)": 432, "CDI (% a.a.)": 4389, "IPCA 12 meses (%)": 13522,
              "IPCA no mês (%)": 433, "Dólar PTAX venda (R$)": 1}
    out = {}
    for nome, cod in series.items():
        try:
            item = sgs_ultimos(cod, 1)[-1]
            out[nome] = (float(item["valor"]), item["data"])
        except Exception as e:  # noqa
            log.warning("SGS %s: %s", cod, e)
    if not out:
        raise ValueError("Banco Central indisponivel")
    return out


def buscar_focus():
    base = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoAnuais"
    anos = [str(HOJE.year), str(HOJE.year + 1)]
    linhas = []
    for ind in ["IPCA", "Selic", "PIB Total", "Câmbio"]:
        filtro = (f"Indicador eq '{ind}' and (DataReferencia eq '{anos[0]}' or "
                  f"DataReferencia eq '{anos[1]}')")
        url = (f"{base}?$top=60&$format=json&$orderby=" + urllib.parse.quote("Data desc")
               + "&$filter=" + urllib.parse.quote(filtro))
        try:
            vals = json.loads(http_get(url, timeout=40)).get("value", [])
        except Exception as e:  # noqa
            log.warning("Focus %s: %s", ind, e)
            continue
        for ano in anos:
            cand = [v for v in vals if str(v.get("DataReferencia")) == ano
                    and v.get("baseCalculo", 0) in (0, "0")]
            if cand:
                v = max(cand, key=lambda x: x["Data"])
                linhas.append({"Indicador": ind, "Ano": ano, "Mediana": v.get("Mediana"),
                               "Data": v.get("Data")})
    if not linhas:
        raise ValueError("Focus indisponivel")
    return pd.DataFrame(linhas)


def yahoo(simbolo):
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/"
           f"{urllib.parse.quote(simbolo)}?range=5d&interval=1d")
    res = json.loads(http_get(url, timeout=15, tentativas=1))["chart"]["result"][0]
    meta = res["meta"]
    fech = [c for c in res["indicators"]["quote"][0].get("close", []) if c is not None]
    preco = meta.get("regularMarketPrice") or (fech[-1] if fech else None)
    ant = fech[-2] if len(fech) >= 2 else meta.get("chartPreviousClose")
    var = (preco / ant - 1) if preco and ant else math.nan
    return {"Preco": preco, "VarDia": var, "Max52": meta.get("fiftyTwoWeekHigh"),
            "Min52": meta.get("fiftyTwoWeekLow")}


def yahoo_lote(simbolos):
    out = {}

    def um(s):
        try:
            return s, yahoo(s)
        except Exception:  # noqa
            return s, None
    with ThreadPoolExecutor(max_workers=8) as ex:
        for s, r in ex.map(um, simbolos):
            if r and r["Preco"]:
                out[s] = r
    if not out:
        raise ValueError("Yahoo Finance indisponivel")
    return out


def buscar_tesouro():
    CACHE.mkdir(exist_ok=True)
    arq = CACHE / "tesouro.csv"
    recente = arq.exists() and dt.date.fromtimestamp(arq.stat().st_mtime) == HOJE
    if not recente or DEMO:
        raiz = ("https://www.tesourotransparente.gov.br/ckan/dataset/"
                "df56aa42-484a-4a59-8184-7676580c81e3/resource/"
                "796d2059-14e9-44e3-80c9-2d9e30b405c1/download/")
        ultimo = None
        for nome in ("PrecoTaxaTesouroDireto.csv", "precotaxatesourodireto.csv"):
            try:
                arq.write_bytes(http_get(raiz + nome, timeout=120))
                ultimo = None
                break
            except Exception as e:  # noqa
                ultimo = e
        if ultimo and not arq.exists():
            raise ultimo
    df = pd.read_csv(arq, sep=";", decimal=",", encoding="latin-1")
    df.columns = [c.strip() for c in df.columns]
    df["Data Base"] = pd.to_datetime(df["Data Base"], format="%d/%m/%Y")
    df["Data Vencimento"] = pd.to_datetime(df["Data Vencimento"], format="%d/%m/%Y")
    df = df[df["Data Base"] == df["Data Base"].max()]
    df = df[df["Data Vencimento"] > pd.Timestamp(HOJE)]
    out = pd.DataFrame({
        "Título": df["Tipo Titulo"], "Vencimento": df["Data Vencimento"].dt.date,
        "Taxa compra (% a.a.)": df["Taxa Compra Manha"].astype(float),
        "PU compra (R$)": df["PU Compra Manha"].astype(float),
        "Data base": df["Data Base"].dt.date})
    return out.sort_values(["Título", "Vencimento"]).reset_index(drop=True)


FORTES = [r"fato relevante", r"dividendos?", r"jcp", r"juros sobre capital", r"proventos?",
          r"recupera[çc][ãa]o judicial", r"opa", r"oferta p[úu]blica", r"emiss[ãa]o de cotas",
          r"subscri[çc][ãa]o", r"follow-on", r"grupamento", r"desdobramento", r"rebaixa\w*",
          r"calote", r"inadimpl\w*", r"fraude", r"investiga[çc][ãa]o", r"fus[ãa]o", r"aquisi[çc][ãa]o",
          r"recompra", r"guidance"]
GERAIS = [r"resultados?", r"balan[çc]o", r"lucro", r"preju[íi]zo", r"vac[âa]ncia", r"rendimentos?",
          r"pre[çc]o-alvo", r"recomenda[çc][ãa]o"]
RE_FORTE = re.compile(r"\b(" + "|".join(FORTES) + r")\b", re.I)
RE_GERAL = re.compile(r"\b(" + "|".join(GERAIS) + r")\b", re.I)


def buscar_noticias(ticker, n):
    q = urllib.parse.quote(f'"{ticker}" when:7d')
    url = f"https://news.google.com/rss/search?q={q}&hl=pt-BR&gl=BR&ceid=BR:pt-419"
    raiz = ET.fromstring(http_get(url, timeout=20, tentativas=1))
    itens = []
    for it in raiz.iter("item"):
        titulo = (it.findtext("title") or "").strip()
        try:
            quando = parsedate_to_datetime(it.findtext("pubDate")).astimezone().replace(tzinfo=None)
        except Exception:  # noqa
            quando = None
        itens.append({"Ticker": ticker, "Data": quando, "Título": titulo,
                      "Fonte": (it.findtext("source") or "").strip(),
                      "Relevante": "ALTA" if RE_FORTE.search(titulo) else "MÉDIA" if RE_GERAL.search(titulo) else "",
                      "Link": (it.findtext("link") or "").strip()})
    itens.sort(key=lambda x: x["Data"] or dt.datetime.min, reverse=True)
    return itens[:n]


# ------------------------------------------------------------------- carteira
CAB_RV = ["Ticker", "Quantidade", "Preço médio (R$)", "Alerta se cair abaixo de (R$)",
          "Alerta se subir acima de (R$)", "Observações"]
CAB_RF = ["Nome do investimento", "Tipo", "Indexador", "Taxa", "Valor aplicado (R$)",
          "Data da aplicação", "Vencimento", "Isento de IR? (S/N)"]
CAB_META = ["Classe", "Alvo (%)"]
CAB_COMPRAS = ["Data", "Ticker", "Tipo (Compra/Venda)", "Quantidade", "Preço unitário (R$)",
               "Taxas e corretagem (R$)", "Observações"]
LEGENDA_COMPRAS = (
    "Registre aqui CADA compra (e cada venda) de ações e FIIs, uma por linha, nas células amarelas. "
    "Data = dia da operação (dd/mm/aaaa). Ticker = código na B3 (ex.: ITUB4). Tipo = Compra ou Venda. "
    "Preço unitário = o preço exato pago por ação ou cota. Taxas e corretagem é opcional. "
    "O robô soma as compras, calcula sozinho a quantidade e o preço médio de cada ativo e mostra o histórico no Dashboard. "
    "Se o ativo está aqui, não precisa preencher Quantidade e Preço médio na aba Renda Variavel (ela serve só para os alertas de preço). "
    "Venda reduz a quantidade e não muda o preço médio. Salve e feche o arquivo; a atualização vem em até 15 min.")


def criar_aba_compras(wb):
    """Adiciona a aba Compras (historico de operacoes). Sem linhas de exemplo, para nao entrar na carteira."""
    from openpyxl.worksheet.datavalidation import DataValidation
    ws = wb.create_sheet("Compras")
    azul = PatternFill("solid", fgColor="1F3864")
    amar = PatternFill("solid", fgColor="FFF2CC")
    ws["A1"] = "MINHAS COMPRAS E VENDAS — AÇÕES e FIIs"
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    for c in range(1, len(CAB_COMPRAS) + 1):
        ws.cell(1, c).fill = azul
    ws["A2"] = LEGENDA_COMPRAS
    ws["A2"].font = Font(name="Arial", size=9, italic=True)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(CAB_COMPRAS))
    ws.row_dimensions[2].height = 92
    for i, (h, w) in enumerate(zip(CAB_COMPRAS, [14, 12, 18, 14, 18, 20, 36]), 1):
        c = ws.cell(4, i, h)
        c.font, c.fill = Font(name="Arial", size=10, bold=True, color="FFFFFF"), azul
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[4].height = 32
    for r in range(5, 505):
        for i in range(1, len(CAB_COMPRAS) + 1):
            c = ws.cell(r, i)
            c.fill, c.font = amar, Font(name="Arial", size=10, color="0000FF")
        ws.cell(r, 1).number_format = "DD/MM/YYYY"
        ws.cell(r, 5).number_format = ws.cell(r, 6).number_format = "#,##0.00"
    dv = DataValidation(type="list", formula1='"Compra,Venda"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("C5:C504")
    ws.freeze_panes = "A5"
    return ws


def garantir_aba_compras(caminho):
    """Se Minha_Carteira.xlsx ainda nao tem a aba Compras, cria (uma vez). Falha em silencio se o arquivo estiver aberto."""
    try:
        wb = load_workbook(caminho)
        if "Compras" in wb.sheetnames:
            wb.close()
            return False
        criar_aba_compras(wb)
        wb.save(caminho)
        log.info("Aba Compras criada em %s", caminho)
        return True
    except Exception as e:  # noqa
        log.warning("Nao foi possivel criar a aba Compras agora (%s); tento na proxima atualizacao", e)
        return False


def criar_modelo_carteira(caminho):
    wb = Workbook()
    f_tit = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    f_cab = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    f_ent = Font(name="Arial", size=10, color="0000FF")
    f_txt = Font(name="Arial", size=10)
    azul = PatternFill("solid", fgColor="1F3864")
    amar = PatternFill("solid", fgColor="FFF2CC")

    def aba(ws, titulo, legenda, cab, exemplos, larguras):
        ws["A1"] = titulo
        ws["A1"].font = f_tit
        for c in range(1, len(cab) + 1):
            ws.cell(1, c).fill = azul
        ws["A2"] = legenda
        ws["A2"].font = Font(name="Arial", size=9, italic=True)
        ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(cab))
        ws.row_dimensions[2].height = 62
        for i, h in enumerate(cab, 1):
            c = ws.cell(4, i, h)
            c.font, c.fill = f_cab, azul
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            ws.column_dimensions[get_column_letter(i)].width = larguras[i - 1]
        ws.row_dimensions[4].height = 32
        for r, linha in enumerate(exemplos, 5):
            for i, v in enumerate(linha, 1):
                c = ws.cell(r, i, v)
                c.font, c.fill = f_ent, amar
                if isinstance(v, dt.date):
                    c.number_format = "DD/MM/YYYY"
        for r in range(5 + len(exemplos), 60):
            for i in range(1, len(cab) + 1):
                ws.cell(r, i).fill = amar
                ws.cell(r, i).font = f_ent
        ws.freeze_panes = "A5"

    ws = wb.active
    ws.title = "Renda Variavel"
    aba(ws, "MINHA CARTEIRA — AÇÕES, FIIs, ETFs e BDRs",
        "Preencha só as células amarelas, uma linha por ativo. As linhas abaixo são EXEMPLOS: "
        "substitua pelos seus ativos. Ticker = código na B3 (ex.: PETR4, MXRF11). Os dois campos "
        "de alerta são opcionais: se o preço cruzar o valor, o robô avisa. Salve e feche o arquivo; "
        "o Painel reflete as mudanças na próxima atualização (até 15 min).",
        CAB_RV,
        [["PETR4", 200, 34.50, 30, 45, "exemplo"], ["ITUB4", 150, 31.20, None, None, "exemplo"],
         ["VALE3", 80, 62.00, 52, None, "exemplo"], ["WEGE3", 60, 38.90, None, None, "exemplo"],
         ["MXRF11", 500, 9.80, 9, None, "exemplo"], ["HGLG11", 30, 158.00, None, None, "exemplo"],
         ["KNCR11", 40, 101.50, None, None, "exemplo"]],
        [14, 13, 16, 20, 20, 34])
    ws = wb.create_sheet("Renda Fixa")
    aba(ws, "MINHA CARTEIRA — RENDA FIXA",
        "Preencha só as células amarelas. Tipo: CDB, LCI, LCA, Tesouro, Debênture, CRI, CRA ou Outro. "
        "Indexador e Taxa: CDI = % do CDI (ex.: 110); SELIC = ágio em % a.a. (ex.: 0,1); "
        "IPCA = juro real em % a.a. (ex.: 7,5 para IPCA+7,5%); PRE = taxa em % a.a. (ex.: 14,5). "
        "As linhas abaixo são EXEMPLOS: substitua pelos seus investimentos. O valor atual é uma "
        "estimativa na curva do papel (não considera marcação a mercado).",
        CAB_RF,
        [["Tesouro Selic 2031", "Tesouro", "SELIC", 0.1, 20000, dt.date(2025, 3, 10), dt.date(2031, 3, 1), "N"],
         ["CDB Banco X 110% CDI", "CDB", "CDI", 110, 15000, dt.date(2025, 8, 5), dt.date(2028, 8, 5), "N"],
         ["Tesouro IPCA+ 2035", "Tesouro", "IPCA", 7.4, 10000, dt.date(2025, 11, 20), dt.date(2035, 5, 15), "N"],
         ["LCI Banco Y 92% CDI", "LCI", "CDI", 92, 12000, dt.date(2026, 2, 2), dt.date(2026, 12, 2), "S"]],
        [30, 13, 12, 10, 18, 16, 14, 16])
    ws = wb.create_sheet("Metas")
    aba(ws, "ALOCAÇÃO-ALVO (opcional)",
        "Informe o percentual que VOCÊ deseja ter em cada classe (soma = 100). O Painel mostra o "
        "desvio da carteira em relação ao seu alvo e o valor necessário para rebalancear. "
        "Deixe em branco para não usar.",
        CAB_META, [["Renda Fixa", 50], ["Ações", 30], ["FIIs", 20]], [22, 14])
    criar_aba_compras(wb)
    wb.save(caminho)


def _ler_aba(wb, nome, chave):
    """Retorna as linhas de dados como listas de CELULAS (valor + formato)."""
    if nome not in wb.sheetnames:
        return []
    linhas = list(wb[nome].iter_rows())
    ini = next((i for i, l in enumerate(linhas) if l and l[0].value and str(l[0].value).strip() == chave), None)
    if ini is None:
        return []
    return [l for l in linhas[ini + 1:] if l and l[0].value not in (None, "")]


def _v(cel):
    return cel.value if hasattr(cel, "value") else cel


def _num_pct(cel):
    """Numero da celula; se a celula estiver formatada como %, o Excel guarda 100% como 1 -> devolve 100."""
    v = _num(_v(cel))
    fmt = getattr(cel, "number_format", "") or ""
    if v == v and "%" in fmt:
        v *= 100
    return v


def _num(v):
    if v is None or v == "":
        return math.nan
    if isinstance(v, (int, float)):
        return float(v)
    return br_num(v) if "," in str(v) else float(str(v))


def _data(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    try:
        return dt.datetime.strptime(str(v).strip(), "%d/%m/%Y").date()
    except Exception:  # noqa
        return None


def ler_carteira(caminho):
    wb = load_workbook(caminho, data_only=True)
    vazio = type("C", (), {"value": None, "number_format": "General"})()
    rv = []
    for l in _ler_aba(wb, "Renda Variavel", "Ticker"):
        l = list(l) + [vazio] * 6
        t = str(_v(l[0])).strip().upper().replace(".SA", "")
        q, pm = _num(_v(l[1])), _num(_v(l[2]))
        if t and q == q and q > 0:
            rv.append({"Ticker": t, "Qtd": q, "PM": pm, "AlertaBaixo": _num(_v(l[3])),
                       "AlertaAlto": _num(_v(l[4]))})
    rf = []
    for l in _ler_aba(wb, "Renda Fixa", "Nome do investimento"):
        l = list(l) + [vazio] * 8
        val = _num(_v(l[4]))
        if val == val and val > 0:
            idx = str(_v(l[2]) or "").strip().upper().replace("É", "E")
            taxa = _num_pct(l[3])
            if idx == "CDI" and taxa == taxa and 0 < taxa <= 3:
                taxa *= 100  # 1 ou 1,2 so faz sentido como 100% ou 120% do CDI
            rf.append({"Nome": str(_v(l[0])).strip(), "Tipo": str(_v(l[1]) or "").strip(),
                       "Indexador": idx, "Taxa": taxa, "Aplicado": val, "DataAplic": _data(_v(l[5])),
                       "Vencimento": _data(_v(l[6])),
                       "Isento": str(_v(l[7]) or "").strip().upper().startswith("S")})
    # ---- aba Compras: calcula quantidade e preco medio por ativo (custo medio) e guarda cada operacao
    lotes = []
    for n_, l in enumerate(_ler_aba(wb, "Compras", "Data")):
        l = list(l) + [vazio] * 7
        d = _data(_v(l[0]))
        t = str(_v(l[1]) or "").strip().upper().replace(".SA", "")
        q, pu, tx = _num(_v(l[3])), _num(_v(l[4])), _num(_v(l[5]))
        venda = str(_v(l[2]) or "").strip().upper().startswith("V")
        if t and q == q and q > 0 and pu == pu and pu > 0 and d is not None:
            lotes.append({"Data": d, "Ticker": t, "Tipo": "Venda" if venda else "Compra", "Qtd": q, "Preco": pu,
                          "Taxas": tx if tx == tx else 0.0, "_ordem": n_})
        elif t or _v(l[3]) not in (None, ""):
            log.warning("Aba Compras: linha ignorada (confira data, ticker, quantidade e preço): %s", t or "?")
    lotes.sort(key=lambda x: (x["Data"], x["_ordem"]))
    pos = {}
    for x in lotes:
        p_ = pos.setdefault(x["Ticker"], {"q": 0.0, "custo": 0.0})
        if x["Tipo"] == "Compra":
            p_["q"] += x["Qtd"]
            p_["custo"] += x["Qtd"] * x["Preco"] + x["Taxas"]
        else:
            vend = min(x["Qtd"], p_["q"])
            if vend < x["Qtd"]:
                log.warning("Aba Compras: venda de %s maior que a quantidade em carteira; ajustada", x["Ticker"])
            if p_["q"] > 0:
                p_["custo"] -= p_["custo"] / p_["q"] * vend
            p_["q"] -= vend
    ja = {p["Ticker"]: p for p in rv}
    for t, p_ in pos.items():
        if p_["q"] > 1e-9:
            base_ = ja.get(t)
            novo = {"Ticker": t, "Qtd": p_["q"], "PM": p_["custo"] / p_["q"],
                    "AlertaBaixo": base_["AlertaBaixo"] if base_ else math.nan,
                    "AlertaAlto": base_["AlertaAlto"] if base_ else math.nan}
            if base_:
                rv[rv.index(base_)] = novo
            else:
                rv.append(novo)
        else:
            rv = [p for p in rv if p["Ticker"] != t]
    metas = {}
    for l in _ler_aba(wb, "Metas", "Classe"):
        if len(l) < 2:
            continue
        v = _num(_v(l[1]))
        if v == v:
            metas[str(_v(l[0])).strip()] = v if "%" in (getattr(l[1], "number_format", "") or "") else v / 100
    wb.close()
    return {"rv": rv, "rf": rf, "metas": metas, "lotes": lotes}


def aliquota_ir(dias):
    return 0.225 if dias <= 180 else 0.20 if dias <= 360 else 0.175 if dias <= 720 else 0.15


def valorizar_rf(rf, series):
    linhas = []
    for p in rf:
        d0 = p["DataAplic"] or HOJE
        dias = max((HOJE - d0).days, 0)
        du = int(np.busday_count(d0, HOJE)) if dias else 0
        taxa = p["Taxa"] if p["Taxa"] == p["Taxa"] else 0.0
        idx, fator, obs = p["Indexador"], 1.0, ""
        try:
            if idx in ("CDI", "SELIC"):
                s = series.get("cdi" if idx == "CDI" else "selic")
                if s is None:
                    raise ValueError("série diária indisponível")
                v = s[(s["data"] >= d0) & (s["data"] < HOJE)]["valor"] / 100
                if idx == "CDI":
                    fator = float(np.prod(1 + v * taxa / 100))
                else:
                    fator = float(np.prod(1 + v)) * (1 + taxa / 100) ** (du / 252)
            elif idx == "IPCA":
                s = series.get("ipca")
                if s is None:
                    raise ValueError("IPCA indisponível")
                ini = d0.replace(day=1)
                v = s[s["data"] >= ini]["valor"] / 100
                fator = float(np.prod(1 + v)) * (1 + taxa / 100) ** (du / 252)
            elif idx == "PRE":
                fator = (1 + taxa / 100) ** (du / 252)
            else:
                obs = "Indexador não reconhecido"
        except Exception as e:  # noqa
            obs = f"Sem cálculo: {e}"
        bruto = p["Aplicado"] * fator
        rend = bruto - p["Aplicado"]
        aliq = 0.0 if p["Isento"] else aliquota_ir(dias)
        liquido = bruto - max(rend, 0) * aliq
        venc = p["Vencimento"]
        dias_venc = (venc - HOJE).days if venc else None
        tx = f"{taxa:g}".replace(".", ",")
        desc = {"CDI": f"{tx}% do CDI", "SELIC": f"Selic + {tx}%",
                "IPCA": f"IPCA + {tx}%", "PRE": f"{tx}% a.a."}.get(idx, idx)
        linhas.append({"Investimento": p["Nome"], "Tipo": p["Tipo"], "Remuneração": desc,
                       "Aplicado (R$)": p["Aplicado"], "Data aplicação": d0,
                       "Valor bruto estimado (R$)": bruto, "Rendimento (R$)": rend,
                       "Rentab. (%)": fator - 1, "IR estimado": aliq,
                       "Valor líquido estimado (R$)": liquido, "Vencimento": venc,
                       "Dias p/ vencer": dias_venc, "Pontos de atenção": obs})
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------ projecao
HORIZONTES = [(1, "1 mês"), (3, "3 meses"), (6, "6 meses"), (12, "12 meses"), (24, "24 meses")] + \
    [(12 * a, f"{12 * a} meses") for a in range(3, 11)]


def _meses(d, n):
    m = d.month - 1 + n
    a, m = d.year + m // 12, m % 12 + 1
    dia = min(d.day, [31, 29 if a % 4 == 0 and (a % 100 or a % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1])
    return dt.date(a, m, dia)


def projetar_rf(cart_rf, rf, cdi_aa, focus):
    """Projeta o valor liquido da renda fixa. Retorna (tabela_focus, tabela_juros_parados, premissas)."""
    if not len(rf) or not cdi_aa:
        return None, None, None
    sel_fim, ipca_aa = None, 0.045
    if focus is not None and len(focus):
        f = focus[(focus["Indicador"] == "Selic") & (focus["Ano"] == str(HOJE.year + 1))]
        if len(f):
            sel_fim = float(f.iloc[0]["Mediana"])
        f = focus[(focus["Indicador"] == "IPCA") & (focus["Ano"] == str(HOJE.year + 1))]
        if len(f):
            ipca_aa = float(f.iloc[0]["Mediana"]) / 100
    cdi0 = cdi_aa / 100
    cdi_fim = (sel_fim - 0.10) / 100 if sel_fim else cdi0
    alvo = dt.date(HOJE.year + 1, 12, 31)
    total_dias = max((alvo - HOJE).days, 1)

    def cdi_em(d, cenario):
        if cenario == "parado":
            return cdi0
        t = min(max((d - HOJE).days / total_dias, 0), 1)
        return cdi0 + (cdi_fim - cdi0) * t

    def valor(p, linha, hor_data, cenario):
        venc = p["Vencimento"]
        fim = min(hor_data, venc) if venc and venc > HOJE else hor_data
        bruto = float(linha["Valor bruto estimado (R$)"])
        taxa = p["Taxa"] if p["Taxa"] == p["Taxa"] else 0.0
        idx = p["Indexador"]
        d = HOJE
        while d < fim:
            if d.weekday() < 5:
                if idx == "CDI":
                    bruto *= 1 + ((1 + cdi_em(d, cenario)) ** (1 / 252) - 1) * taxa / 100
                elif idx == "SELIC":
                    bruto *= (1 + cdi_em(d, cenario) + 0.001) ** (1 / 252) * (1 + taxa / 100) ** (1 / 252)
                elif idx == "IPCA":
                    bruto *= ((1 + ipca_aa) * (1 + taxa / 100)) ** (1 / 252)
                elif idx == "PRE":
                    bruto *= (1 + taxa / 100) ** (1 / 252)
            d += dt.timedelta(days=1)
        d0 = p["DataAplic"] or HOJE
        rend = bruto - p["Aplicado"]
        aliq = 0.0 if p["Isento"] else aliquota_ir(max((fim - d0).days, 0))
        return bruto - max(rend, 0) * aliq

    out = {}
    for cen in ("focus", "parado"):
        linhas = []
        for p, (_, l) in zip(cart_rf, rf.iterrows()):
            lin = {"Investimento": p["Nome"], "Remuneração": l["Remuneração"],
                   "Hoje (líquido, R$)": l["Valor líquido estimado (R$)"]}
            for n, rot in HORIZONTES:
                lin[rot] = valor(p, l, _meses(HOJE, n), cen)
            linhas.append(lin)
        df = pd.DataFrame(linhas)
        tot = {c: (df[c].sum() if c not in ("Investimento", "Remuneração") else "") for c in df.columns}
        tot["Investimento"] = "TOTAL"
        df = pd.concat([df, pd.DataFrame([tot])], ignore_index=True)
        out[cen] = df
    prem = {"cdi0": cdi0, "cdi_fim": cdi_fim, "sel_fim": sel_fim, "ipca": ipca_aa, "alvo": alvo}
    return out["focus"], out["parado"], prem


# -------------------------------------------------------------------- ranking
def _pct(s, maior_melhor=True):
    r = s.rank(pct=True) * 100
    return r if maior_melhor else 100 - r + (100 / max(s.notna().sum(), 1))


def pontuar_acoes(df, cfg):
    d = df.copy()
    pos = lambda c: d[c].where(d[c] > 0)  # noqa
    d["EarningsYield"] = 1 / pos("PL")
    dy = d["DY"].clip(upper=0.20)
    grupos = {
        "Valor": pd.concat([_pct(d["EarningsYield"]), _pct(pos("PVP"), False),
                            _pct(pos("EV_EBIT"), False)], axis=1).mean(axis=1),
        "Renda": _pct(dy.where(dy > 0)).fillna(0),
        "Qualidade": pd.concat([_pct(d["ROE"]), _pct(d["ROIC"].where(d["ROIC"] != 0)),
                                _pct(d["MrgLiq"].where(d["MrgLiq"] != 0)),
                                _pct(d["DivLiq_Patrim"], False)], axis=1).mean(axis=1),
        "Crescimento": _pct(d["Cresc5a"].where(d["Cresc5a"] != 0)),
    }
    pesos = {"Valor": cfg_f(cfg, "pesos_acoes", "valor", 30), "Renda": cfg_f(cfg, "pesos_acoes", "renda", 20),
             "Qualidade": cfg_f(cfg, "pesos_acoes", "qualidade", 35),
             "Crescimento": cfg_f(cfg, "pesos_acoes", "crescimento", 15)}
    num = sum(grupos[g].fillna(0) * pesos[g] for g in grupos)
    den = sum(grupos[g].notna() * pesos[g] for g in grupos)
    d["Score"] = (num / den.replace(0, np.nan)).round(1)
    d.loc[~(d["PL"] > 0), "Score"] = np.nan  # empresas com prejuízo ficam fora do ranking
    for g in grupos:
        d["Nota " + g] = grupos[g].round(0)
    lpa, vpa = d["Preco"] / pos("PL"), d["Preco"] / pos("PVP")
    d["Graham"] = np.sqrt(22.5 * lpa * vpa)
    d["MargemGraham"] = d["Graham"] / d["Preco"] - 1
    d["Bazin"] = d["DY"] * d["Preco"] / 0.06
    return d


def pontuar_fiis(df, cfg):
    d = df.copy()
    dy = d["DY"].where((d["DY"] > 0) & (d["DY"] <= 0.20))
    pvp = d["PVP"].where(d["PVP"] > 0.4)
    grupos = {
        "Renda": _pct(dy),
        "Preço": _pct((pvp - 0.95).abs().where(pvp.notna()), False),
        "Liquidez": _pct(d["Liquidez"]),
        "Vacância": _pct(d["Vacancia"], False),
    }
    pesos = {"Renda": cfg_f(cfg, "pesos_fiis", "renda", 40), "Preço": cfg_f(cfg, "pesos_fiis", "preco", 30),
             "Liquidez": cfg_f(cfg, "pesos_fiis", "liquidez", 15),
             "Vacância": cfg_f(cfg, "pesos_fiis", "vacancia", 15)}
    num = sum(grupos[g].fillna(0) * pesos[g] for g in grupos)
    den = sum(grupos[g].notna() * pesos[g] for g in grupos)
    d["Score"] = (num / den.replace(0, np.nan)).round(1)
    d.loc[dy.isna() | pvp.isna(), "Score"] = np.nan
    for g in grupos:
        d["Nota " + g] = grupos[g].round(0)
    return d


def motivo(linha, grupos):
    notas = sorted(((linha.get("Nota " + g), g) for g in grupos
                    if linha.get("Nota " + g) == linha.get("Nota " + g)), reverse=True)
    fortes = [f"{g} ({n:.0f})" for n, g in notas[:2]]
    fracos = [f"{g} ({n:.0f})" for n, g in notas[-1:] if n < 40]
    txt = "Destaques: " + ", ".join(fortes)
    return txt + ("; ponto fraco: " + ", ".join(fracos) if fracos else "")


# --------------------------------------------------------------------- sinais
def sinais_rv(l, cfg, total):
    s = []
    conc = cfg_f(cfg, "sinais", "concentracao_max", 20) / 100
    if total and l["Valor"] / total > conc:
        s.append(f"Concentração de {fpc(l['Valor'] / total, 0)} da carteira (limite {fpc(conc, 0)})")
    res = l.get("ResultadoPct")
    if res == res and res <= cfg_f(cfg, "sinais", "queda_vs_pm", -15) / 100:
        s.append(f"{fpc(abs(res), 0)} abaixo do seu preço médio")
    if res == res and res >= cfg_f(cfg, "sinais", "alta_vs_pm", 30) / 100:
        s.append(f"{fpc(res, 0)} acima do seu preço médio")
    if l["Classe"] == "Ações":
        if l.get("PL") == l.get("PL") and l.get("PL") <= 0:
            s.append("Empresa com prejuízo nos últimos 12 meses (P/L negativo)")
        if l.get("DivLiq_Patrim") == l.get("DivLiq_Patrim") and l["DivLiq_Patrim"] > 2:
            s.append("Dívida líquida acima de 2x o patrimônio")
    if l["Classe"] == "FIIs":
        if l.get("PVP") == l.get("PVP") and l["PVP"] > cfg_f(cfg, "sinais", "fii_pvp_alto", 1.15):
            s.append(f"Cota a {fbr(l['PVP'])}x o valor patrimonial")
        if l.get("Vacancia") == l.get("Vacancia") and l["Vacancia"] > 0.15:
            s.append(f"Vacância de {fpc(l['Vacancia'], 0)}")
        cdi = l.get("_cdi")
        if cdi and l.get("DY") == l.get("DY") and 0 < l["DY"] < cdi * 0.85 * 0.8:
            s.append("Dividend yield abaixo do CDI líquido")
    pos = l.get("PosicaoRanking")
    if pos == pos and pos is not None:
        if pos <= 0.25:
            s.append("No quartil inferior do ranking por critérios")
        elif pos >= 0.75:
            s.append("No quartil superior do ranking por critérios")
    if l.get("Preco") != l.get("Preco"):
        s.append("Cotação não encontrada: confira o ticker")
    return "; ".join(s) if s else "Sem pontos de atenção pelos critérios"


# ------------------------------------------------------- alertas automaticos de preco
OBS_AUTO = "Alertas sugeridos pelo robô (edite à vontade)"


def sugerir_alertas(classe, pm, preco, min52, max52, pvp, cfg):
    """Sugestao mecanica de limites a partir do preco medio e da oscilacao de 52 semanas.
    Distancia = 35% da amplitude de 52 semanas, entre um piso e um teto por classe.
    Alerta de queda = PM - distancia. Alerta de alta = PM + 2x distancia (relacao ganho/risco de 2 para 1);
    em FIIs, limitado ao valor patrimonial x P/VP maximo dos sinais. Nao e recomendacao."""
    fii = classe == "FIIs"
    piso = cfg_f(cfg, "alertas_auto", "queda_min_fii" if fii else "queda_min_acao", 5 if fii else 8) / 100
    teto = cfg_f(cfg, "alertas_auto", "queda_max_fii" if fii else "queda_max_acao", 10 if fii else 20) / 100
    padrao = cfg_f(cfg, "alertas_auto", "queda_padrao_fii" if fii else "queda_padrao_acao", 7 if fii else 12) / 100
    amp = (max52 - min52) / preco if all(x == x and x for x in (min52, max52, preco)) and max52 > min52 else None
    dist = min(teto, max(piso, 0.35 * amp)) if amp else padrao
    ganho = cfg_f(cfg, "alertas_auto", "relacao_ganho_risco", 2.0)
    baixo = pm * (1 - dist)
    alto = pm * (1 + ganho * dist)
    if fii and pvp == pvp and pvp and preco == preco and preco:
        teto_vp = preco / pvp * cfg_f(cfg, "sinais", "fii_pvp_alto", 1.15)
        if pm * 1.03 <= teto_vp < alto:
            alto = teto_vp
    return round(baixo, 2), round(alto, 2)


def sincronizar_carteira(arq, cart, acoes, fiis, cot, cfg):
    """Mantem a aba Renda Variavel de Minha_Carteira.xlsx em dia: posicoes vindas da aba Compras e
    alertas sugeridos. Nunca sobrescreve um alerta que voce digitou; so mexe nos vazios e nos que o proprio
    robo preencheu antes (e que continuam iguais ao valor sugerido)."""
    if cfg.get("alertas_auto", "ativo", fallback="sim").strip().lower() not in ("sim", "s", "true", "1", "yes"):
        return
    if not cart["rv"]:
        return
    if arq.with_name("~$" + arq.name).exists():
        log.info("Minha_Carteira.xlsx aberta no Excel: alertas automáticos ficam para a próxima atualização")
        return
    try:
        mtime0 = arq.stat().st_mtime
        if cart.get("_mtime") is not None and mtime0 != cart["_mtime"]:
            log.info("Minha_Carteira.xlsx mudou desde a leitura: sincronização fica para a próxima rodada")
            return
        wb = load_workbook(arq)
        ws = wb["Renda Variavel"]
        hdr = next((c.row for c in ws["A"] if c.value and str(c.value).strip() == "Ticker"), None)
        if hdr is None:
            return
        linhas = {}
        vazias = []
        for r in range(hdr + 1, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if v in (None, ""):
                vazias.append(r)
            else:
                linhas[str(v).strip().upper().replace(".SA", "")] = r
        est = carregar_estado()
        auto = est.setdefault("alertas_auto", {})
        mudou = False
        novas = []
        for p in cart["rv"]:
            t, pm, qtd = p["Ticker"], p["PM"], p["Qtd"]
            if t not in linhas:
                if not vazias:
                    continue
                linhas[t] = vazias.pop(0)
                ws.cell(linhas[t], 1, t)
                mudou = True
            r = linhas[t]
            if abs((_num(ws.cell(r, 2).value) or 0) - qtd) > 1e-9:
                ws.cell(r, 2, qtd)
                mudou = True
            if pm == pm and not (abs((_num(ws.cell(r, 3).value) or 0) - pm) < 0.005):
                ws.cell(r, 3, round(pm, 4))
                mudou = True
            if pm != pm:
                continue
            classe, pvp = "Ações", math.nan
            if len(fiis) and t in set(fiis["Ticker"]):
                classe, pvp = "FIIs", fiis[fiis["Ticker"] == t].iloc[0].get("PVP", math.nan)
            y = cot.get(t + ".SA") or {}
            preco = y.get("Preco") or math.nan
            baixo, alto = sugerir_alertas(classe, pm, preco, y.get("Min52") or math.nan,
                                          y.get("Max52") or math.nan, pvp, cfg)
            ant = auto.get(t) or {}
            novo_pm = abs(ant.get("pm", -1) - pm) > 0.005
            for col, chave, val in ((4, "baixo", baixo), (5, "alto", alto)):
                atual = ws.cell(r, col).value
                vazio = atual is None or str(atual).strip() == ""
                igual_auto = (not vazio and chave in ant and isinstance(atual, (int, float))
                              and abs(float(atual) - ant[chave]) < 0.006)
                if vazio or (igual_auto and novo_pm):
                    ws.cell(r, col, val)
                    mudou = True
                    p["AlertaBaixo" if col == 4 else "AlertaAlto"] = val
                    ant[chave] = val
                    if vazio:
                        novas.append(t)
            ant["pm"] = pm
            auto[t] = ant
            if ws.cell(r, 6).value in (None, ""):
                ws.cell(r, 6, OBS_AUTO)
                mudou = True
        if mudou:
            if arq.stat().st_mtime != mtime0:
                log.info("Minha_Carteira.xlsx mudou durante a leitura: alertas automáticos na próxima rodada")
                return
            tmp = arq.with_name("~tmp_" + arq.name)
            wb.save(tmp)
            os.replace(tmp, arq)
            salvar_estado(est)
            log.info("Minha_Carteira.xlsx: Renda Variavel atualizada%s",
                     f" (alertas sugeridos para {', '.join(sorted(set(novas)))})" if novas else "")
    except Exception as e:  # noqa
        log.warning("Alertas automáticos não aplicados: %s", e)


# ------------------------------------------------------------------- alertas
def carregar_estado():
    arq = BASE / "estado.json"
    if arq.exists():
        try:
            return json.loads(arq.read_text(encoding="utf-8"))
        except Exception:  # noqa
            pass
    return {"enviados": {}, "selic": None}


def salvar_estado(est):
    limite = (HOJE - dt.timedelta(days=10)).isoformat()
    est["enviados"] = {k: v for k, v in est["enviados"].items() if v >= limite}
    (BASE / "estado.json").write_text(json.dumps(est, ensure_ascii=False, indent=1), encoding="utf-8")


def enviar_email(cfg, assunto, corpo):
    if not cfg_sim(cfg, "email", "ativo"):
        return "desligado"
    rem, senha = cfg.get("email", "remetente", fallback=""), cfg.get("email", "senha_app", fallback="")
    dest = cfg.get("email", "destinatario", fallback="")
    if not (rem and senha and dest):
        return "configuração incompleta"
    msg = EmailMessage()
    msg["Subject"], msg["From"], msg["To"] = assunto, rem, dest
    msg.set_content(corpo)
    with smtplib.SMTP(cfg.get("email", "smtp_servidor", fallback="smtp.gmail.com"),
                      int(cfg_f(cfg, "email", "smtp_porta", 587)), timeout=30) as s:
        s.starttls()
        s.login(rem, senha.replace(" ", ""))
        s.send_message(msg)
    return "enviado"


def enviar_whatsapp(cfg, texto):
    if not cfg_sim(cfg, "whatsapp", "ativo"):
        return "desligado"
    fone, chave = cfg.get("whatsapp", "telefone", fallback=""), cfg.get("whatsapp", "apikey", fallback="")
    if not (fone and chave):
        return "configuração incompleta"
    url = ("https://api.callmebot.com/whatsapp.php?phone=" + urllib.parse.quote(fone)
           + "&text=" + urllib.parse.quote(texto[:1200]) + "&apikey=" + urllib.parse.quote(chave))
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return "enviado"


def despachar(cfg, alertas):
    """alertas: lista de (chave, nivel, texto). Envia só os ainda não enviados."""
    est = carregar_estado()
    novos = [a for a in alertas if a[0] not in est["enviados"]]
    resultado = {}
    if novos and not DEMO:
        linhas = [f"• {t}" for _, _, t in novos]
        corpo = (f"Alertas da carteira — {AGORA:%d/%m/%Y %H:%M}\n\n" + "\n".join(linhas)
                 + "\n\nAbra o Painel_Investimentos.xlsx para os detalhes.\n"
                   "Mensagem automática; não é recomendação de investimento.")
        for nome, fn in (("e-mail", lambda: enviar_email(cfg, f"[Carteira] {len(novos)} alerta(s)", corpo)),
                         ("WhatsApp", lambda: enviar_whatsapp(cfg, corpo))):
            try:
                resultado[nome] = fn()
            except Exception as e:  # noqa
                resultado[nome] = f"erro: {str(e)[:80]}"
                log.warning("Envio %s falhou: %s", nome, e)
        if any(v == "enviado" for v in resultado.values()):
            for ch, _, _ in novos:
                est["enviados"][ch] = HOJE.isoformat()
    salvar_estado(est)
    return novos, resultado


# --------------------------------------------------------------------- Excel
ARIAL = "Arial"
F_N = Font(name=ARIAL, size=10)
F_B = Font(name=ARIAL, size=10, bold=True)
F_CAB = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
F_TIT = Font(name=ARIAL, size=15, bold=True, color="FFFFFF")
F_SEC = Font(name=ARIAL, size=11, bold=True, color="1F3864")
F_NOTA = Font(name=ARIAL, size=9, italic=True, color="595959")
F_POS = Font(name=ARIAL, size=10, color="00803C")
F_NEG = Font(name=ARIAL, size=10, color="C00000")
F_LINK = Font(name=ARIAL, size=10, color="0563C1", underline="single")
AZUL = PatternFill("solid", fgColor="1F3864")
CINZA = PatternFill("solid", fgColor="F2F2F2")
VERDE_F = PatternFill("solid", fgColor="E2EFDA")
VERM_F = PatternFill("solid", fgColor="FCE4D6")
AMAR_F = PatternFill("solid", fgColor="FFF2CC")
BORDA = Border(bottom=Side(style="thin", color="D9D9D9"))
RS, PC, N2, N0, DT = 'R$ #,##0.00', '0.00%', '#,##0.00', '#,##0', 'DD/MM/YYYY'
AVISO = ("Rankings e sinais são cálculos objetivos sobre dados públicos (Fundamentus, Yahoo Finance, "
         "Banco Central, Tesouro Nacional, Google Notícias). Não constituem recomendação de "
         "investimento. Cotações com atraso de 15 minutos ou mais.")


def _limpo(v):
    if v is None:
        return None
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    if isinstance(v, (np.floating,)):
        return None if np.isnan(v) else float(v)
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    return v


def titulo(ws, texto, ncols, sub=None):
    ws["A1"] = texto
    ws["A1"].font = F_TIT
    ws.row_dimensions[1].height = 28
    for c in range(1, ncols + 1):
        ws.cell(1, c).fill = AZUL
    ws["A2"] = sub or f"Atualizado em {AGORA:%d/%m/%Y às %H:%M}"
    ws["A2"].font = F_NOTA
    ws.sheet_view.showGridLines = False


def tabela(ws, r0, df, fmts=None, cores=None, larguras=None, link=None, filtro=True, destaque=None):
    """Escreve df a partir da linha r0. Retorna a linha seguinte livre."""
    fmts, cores = fmts or {}, cores or []
    cols = [c for c in df.columns if c != link]
    for j, c in enumerate(cols, 1):
        cel = ws.cell(r0, j, c)
        cel.font, cel.fill = F_CAB, AZUL
        cel.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        larg = (larguras or {}).get(c)
        if larg:
            ws.column_dimensions[get_column_letter(j)].width = larg
        elif ws.column_dimensions[get_column_letter(j)].width < 12:
            ws.column_dimensions[get_column_letter(j)].width = max(12, min(len(str(c)) + 3, 22))
    ws.row_dimensions[r0].height = 30
    for i, (_, lin) in enumerate(df.iterrows(), 1):
        for j, c in enumerate(cols, 1):
            v = _limpo(lin[c])
            cel = ws.cell(r0 + i, j, v)
            cel.font, cel.border = F_N, BORDA
            if c in fmts:
                cel.number_format = fmts[c]
            if c in cores and isinstance(v, (int, float)):
                cel.font = F_POS if v > 0 else F_NEG if v < 0 else F_N
            if link and c == "Título" and lin.get(link):
                cel.hyperlink, cel.font = lin[link], F_LINK
            if destaque and c == destaque[0]:
                f = destaque[1](v)
                if f:
                    cel.fill = f
            if isinstance(v, str) and len(v) > 40:
                cel.alignment = Alignment(wrap_text=True, vertical="top")
    if filtro and len(df):
        ws.auto_filter.ref = f"A{r0}:{get_column_letter(len(cols))}{r0 + len(df)}"
    return r0 + len(df) + 2


def secao(ws, r, texto):
    ws.cell(r, 1, texto).font = F_SEC
    return r + 1


def par(ws, r, rotulo, valor, fmt=None, nota=None, cor=False):
    ws.cell(r, 1, rotulo).font = F_N
    c = ws.cell(r, 2, _limpo(valor))
    c.font = F_B
    if fmt:
        c.number_format = fmt
    if cor and isinstance(valor, (int, float)) and valor == valor:
        c.font = Font(name=ARIAL, size=10, bold=True, color="00803C" if valor >= 0 else "C00000")
    if nota:
        ws.cell(r, 3, nota).font = F_NOTA
    return r + 1


def salvar(wb, destino):
    tmp = destino.with_name("~tmp_" + destino.name)
    wb.save(tmp)
    try:
        os.replace(tmp, destino)
        return destino
    except PermissionError:
        alt = destino.with_name(destino.stem + " (nova versao).xlsx")
        os.replace(tmp, alt)
        log.warning("Painel aberto no Excel; gravado em %s", alt.name)
        return alt


# ----------------------------------------------------------------- principal
def main():
    configurar_log()
    cfg = carregar_config()
    log.info("Inicio%s", " (DEMO)" if DEMO else "")

    if "--teste-alerta" in sys.argv:
        txt = f"Teste do robô de investimentos em {AGORA:%d/%m/%Y %H:%M}. Canal funcionando."
        for nome, fn in (("e-mail", lambda: enviar_email(cfg, "[Carteira] Teste de alerta", txt)),
                         ("WhatsApp", lambda: enviar_whatsapp(cfg, txt))):
            try:
                print(nome, "->", fn())
            except Exception as e:  # noqa
                print(nome, "-> ERRO:", e)
        return

    pasta = Path(os.environ.get("ROBO_PASTA_SAIDA") or cfg.get("geral", "pasta_saida", fallback=str(BASE / "saida")))
    aviso_pasta = ""
    if DEMO:
        pasta = BASE / "saida_demo"
    try:
        pasta.mkdir(parents=True, exist_ok=True)
    except Exception:  # noqa
        aviso_pasta = f"Pasta {pasta} indisponível; arquivos gravados em {BASE / 'saida'}"
        pasta = BASE / "saida"
        pasta.mkdir(parents=True, exist_ok=True)
        log.warning(aviso_pasta)
    arq_cart = pasta / cfg.get("geral", "arquivo_carteira", fallback="Minha_Carteira.xlsx")
    arq_painel = pasta / cfg.get("geral", "arquivo_painel", fallback="Painel_Investimentos.xlsx")
    if not arq_cart.exists():
        criar_modelo_carteira(arq_cart)
        log.info("Modelo de carteira criado em %s", arq_cart)
    else:
        garantir_aba_compras(arq_cart)
    try:
        mt_cart = arq_cart.stat().st_mtime
        cart = ler_carteira(arq_cart)
        cart["_mtime"] = mt_cart
        STATUS.append(("carteira", "OK", f"{len(cart['rv'])} ativos de renda variável, {len(cart['rf'])} de renda fixa"))
    except Exception as e:  # noqa
        log.error("Erro ao ler carteira: %s", e)
        STATUS.append(("carteira", "FALHA", str(e)[:120]))
        cart = {"rv": [], "rf": [], "metas": {}, "lotes": []}

    # ---- coleta
    acoes_all = com_cache("acoes_fundamentus", buscar_acoes)
    fiis_all = com_cache("fiis_fundamentus", buscar_fiis)
    macro = com_cache("banco_central", buscar_macro) or {}
    focus = com_cache("focus", buscar_focus)
    tesouro = com_cache("tesouro_direto", buscar_tesouro)
    if acoes_all is None:
        acoes_all = pd.DataFrame(columns=COLS_ACOES)
    if fiis_all is None:
        fiis_all = pd.DataFrame(columns=COLS_FIIS)

    meus = [p["Ticker"] for p in cart["rv"]]
    top_a, top_f = int(cfg_f(cfg, "geral", "top_acoes", 150)), int(cfg_f(cfg, "geral", "top_fiis", 100))
    acoes = acoes_all.sort_values("Liquidez", ascending=False).head(top_a)
    acoes = pd.concat([acoes, acoes_all[acoes_all["Ticker"].isin(meus) & ~acoes_all["Ticker"].isin(acoes["Ticker"])]])
    fiis = fiis_all.sort_values("Liquidez", ascending=False).head(top_f)
    fiis = pd.concat([fiis, fiis_all[fiis_all["Ticker"].isin(meus) & ~fiis_all["Ticker"].isin(fiis["Ticker"])]])
    acoes, fiis = acoes.reset_index(drop=True), fiis.reset_index(drop=True)

    simbolos = ["^BVSP", "BRL=X", "IFIX.SA"] + [t + ".SA" for t in meus]
    if cfg_sim(cfg, "geral", "cotacao_yahoo_mercado"):
        simbolos += [t + ".SA" for t in list(acoes["Ticker"]) + list(fiis["Ticker"])]
    cot = com_cache("yahoo_finance", lambda: yahoo_lote(list(dict.fromkeys(simbolos)))) or {}

    for df in (acoes, fiis):
        y = df["Ticker"].map(lambda t: cot.get(t + ".SA") or {})
        df["VarDia"] = y.map(lambda d: d.get("VarDia", math.nan))
        df["Min52"] = y.map(lambda d: d.get("Min52") or math.nan)
        df["Max52"] = y.map(lambda d: d.get("Max52") or math.nan)
        py = y.map(lambda d: d.get("Preco") or math.nan)
        df["Preco"] = py.where(py.notna(), df["Preco"])
    acoes = pontuar_acoes(acoes, cfg) if len(acoes) else acoes
    fiis = pontuar_fiis(fiis, cfg) if len(fiis) else fiis

    # ---- séries para renda fixa da carteira
    series = {}
    datas = [p["DataAplic"] for p in cart["rf"] if p["DataAplic"]]
    if datas:
        d0 = max(min(datas), HOJE - dt.timedelta(days=3600))
        for nome, cod in (("cdi", 12), ("selic", 11), ("ipca", 433)):
            if any(p["Indexador"] == nome.upper() for p in cart["rf"]):
                try:
                    series[nome] = sgs_periodo(cod, d0.replace(day=1) if nome == "ipca" else d0, HOJE)
                except Exception as e:  # noqa
                    log.warning("Série %s: %s", nome, e)
    rf = valorizar_rf(cart["rf"], series)

    # ---- aba Renda Variavel em dia + alertas sugeridos (so preenche vazios)
    sincronizar_carteira(arq_cart, cart, acoes, fiis, cot, cfg)

    # ---- carteira de renda variável
    cdi_aa = (macro.get("CDI (% a.a.)") or macro.get("Selic meta (% a.a.)") or (None,))[0]
    cdi_frac = cdi_aa / 100 if cdi_aa else None
    linhas = []
    for p in cart["rv"]:
        t = p["Ticker"]
        base, classe, grupos = None, "Outros", []
        if len(fiis) and t in set(fiis["Ticker"]):
            base, classe, tab = fiis[fiis["Ticker"] == t].iloc[0], "FIIs", fiis
            grupos = ["Renda", "Preço", "Liquidez", "Vacância"]
        elif len(acoes) and t in set(acoes["Ticker"]):
            base, classe, tab = acoes[acoes["Ticker"] == t].iloc[0], "Ações", acoes
            grupos = ["Valor", "Renda", "Qualidade", "Crescimento"]
        y = cot.get(t + ".SA") or {}
        preco = y.get("Preco") or (base["Preco"] if base is not None else math.nan)
        valor = p["Qtd"] * preco if preco == preco else p["Qtd"] * (p["PM"] if p["PM"] == p["PM"] else 0)
        custo = p["Qtd"] * p["PM"] if p["PM"] == p["PM"] else math.nan
        l = {"Ticker": t, "Classe": classe, "Qtd": p["Qtd"], "PM": p["PM"], "Preco": preco,
             "VarDia": y.get("VarDia", math.nan), "Valor": valor, "Custo": custo,
             "Resultado": valor - custo if custo == custo else math.nan,
             "ResultadoPct": (valor / custo - 1) if custo == custo and custo else math.nan,
             "AlertaBaixo": p["AlertaBaixo"], "AlertaAlto": p["AlertaAlto"], "_cdi": cdi_frac}
        if base is not None:
            for c in ("DY", "PVP", "PL", "ROE", "DivLiq_Patrim", "Vacancia", "Score", "Graham",
                      "MargemGraham", "Bazin", "Segmento"):
                l[c] = base.get(c, math.nan)
            sc = tab["Score"].dropna()
            l["PosicaoRanking"] = (sc < base["Score"]).mean() if base["Score"] == base["Score"] and len(sc) else math.nan
            l["Motivo"] = motivo(base, grupos) if base["Score"] == base["Score"] else "Fora do ranking (dados insuficientes ou prejuízo)"
        linhas.append(l)
    rv = pd.DataFrame(linhas)
    total_rv = float(rv["Valor"].sum()) if len(rv) else 0.0
    total_rf = float(rf["Valor bruto estimado (R$)"].sum()) if len(rf) else 0.0
    total = total_rv + total_rf
    if len(rv):
        rv["Peso"] = rv["Valor"] / total if total else math.nan
        rv["Sinais"] = rv.apply(lambda l: sinais_rv(l, cfg, total), axis=1)
        ant = (rv["Valor"] / (1 + rv["VarDia"])).where(rv["VarDia"].notna(), rv["Valor"])
        var_cart = total_rv / ant.sum() - 1 if ant.sum() else math.nan
    else:
        var_cart = math.nan

    # ---- notícias
    noticias = []
    n_not = int(cfg_f(cfg, "geral", "noticias_por_ativo", 4))
    falhas = 0
    for t in meus:
        try:
            noticias += buscar_noticias(t, n_not)
        except Exception as e:  # noqa
            falhas += 1
            log.warning("Notícias %s: %s", t, e)
    if meus:
        STATUS.append(("noticias", "OK" if falhas == 0 else "PARCIAL" if falhas < len(meus) else "FALHA",
                       f"{len(noticias)} notícias; {falhas} ativo(s) sem retorno"))
    noticias = pd.DataFrame(noticias, columns=["Ticker", "Data", "Título", "Fonte", "Relevante", "Link"])

    # ---- alertas
    alertas = []
    dia = HOJE.isoformat()
    lim_a, lim_f = cfg_f(cfg, "alertas", "variacao_dia_acao", 4) / 100, cfg_f(cfg, "alertas", "variacao_dia_fii", 2.5) / 100
    for _, l in rv.iterrows() if len(rv) else []:
        v, lim = l["VarDia"], (lim_f if l["Classe"] == "FIIs" else lim_a)
        if v == v and abs(v) >= lim:
            alertas.append((f"var|{l['Ticker']}|{dia}|{'+' if v > 0 else '-'}", "ALTO",
                            f"{l['Ticker']} {fpc(v, 1, True)} no dia (R$ {fbr(l['Preco'])})"))
        if l["AlertaBaixo"] == l["AlertaBaixo"] and l["Preco"] == l["Preco"] and l["Preco"] <= l["AlertaBaixo"]:
            alertas.append((f"baixo|{l['Ticker']}|{dia}", "ALTO",
                            f"{l['Ticker']} a R$ {fbr(l['Preco'])}, abaixo do seu limite de R$ {fbr(l['AlertaBaixo'])}"))
        if l["AlertaAlto"] == l["AlertaAlto"] and l["Preco"] == l["Preco"] and l["Preco"] >= l["AlertaAlto"]:
            alertas.append((f"alto|{l['Ticker']}|{dia}", "ALTO",
                            f"{l['Ticker']} a R$ {fbr(l['Preco'])}, acima do seu limite de R$ {fbr(l['AlertaAlto'])}"))
    if var_cart == var_cart and abs(var_cart) >= cfg_f(cfg, "alertas", "variacao_dia_carteira", 2) / 100:
        alertas.append((f"carteira|{dia}", "ALTO", f"Carteira de renda variável {fpc(var_cart, 1, True)} no dia"))
    for sim, nome, chave in (("^BVSP", "Ibovespa", "variacao_ibovespa"), ("BRL=X", "Dólar", "variacao_dolar")):
        v = (cot.get(sim) or {}).get("VarDia", math.nan)
        if v == v and abs(v) >= cfg_f(cfg, "alertas", chave, 2.5) / 100:
            alertas.append((f"{sim}|{dia}", "MÉDIO", f"{nome} {fpc(v, 1, True)} no dia"))
    est = carregar_estado()
    selic = (macro.get("Selic meta (% a.a.)") or (None,))[0]
    if selic is not None:
        if est.get("selic") not in (None, selic):
            alertas.append((f"selic|{selic}", "ALTO", f"Selic meta mudou de {fbr(est['selic'])}% para {fbr(selic)}% a.a."))
        est["selic"] = selic
        salvar_estado(est)
    por_ticker = {}
    for _, l in noticias[noticias["Relevante"] == "ALTA"].iterrows():
        por_ticker[l["Ticker"]] = por_ticker.get(l["Ticker"], 0) + 1
        if por_ticker[l["Ticker"]] > 2:
            continue
        h = hashlib.md5((l["Link"] or l["Título"]).encode("utf-8")).hexdigest()[:12]
        alertas.append((f"news|{h}", "MÉDIO", f"Notícia {l['Ticker']}: {l['Título'][:140]}"))
    for _, l in rf.iterrows() if len(rf) else []:
        d = l["Dias p/ vencer"]
        if d is not None and d == d and 0 <= d <= cfg_f(cfg, "alertas", "vencimento_rf_dias", 30):
            alertas.append((f"venc|{l['Investimento']}|{HOJE:%Y-%m}", "MÉDIO",
                            f"{l['Investimento']} vence em {int(d)} dia(s) ({l['Vencimento']:%d/%m/%Y})"))
    novos, envio = despachar(cfg, alertas)
    chaves_novas = {a[0] for a in novos}
    enviou = any(v == "enviado" for v in envio.values())

    # ================================================================ planilha
    wb = Workbook()
    # ---------------------------------------------------------------- Painel
    ws = wb.active
    ws.title = "Painel"
    titulo(ws, "PAINEL DE INVESTIMENTOS" + (" — DEMONSTRAÇÃO (dados fictícios)" if DEMO else ""), 6)
    for col, w in zip("ABCDEF", (46, 20, 44, 18, 18, 18)):
        ws.column_dimensions[col].width = w
    r = 4
    r = secao(ws, r, "MINHA CARTEIRA")
    r = par(ws, r, "Patrimônio total", total, RS)
    r = par(ws, r, "Renda variável", total_rv, RS, f"{fpc(total_rv / total)} da carteira" if total else None)
    r = par(ws, r, "Renda fixa (valor bruto estimado)", total_rf, RS, f"{fpc(total_rf / total)} da carteira" if total else None)
    r = par(ws, r, "Variação da renda variável no dia", var_cart, PC, cor=True)
    if len(rv):
        res = rv["Resultado"].sum(skipna=True)
        r = par(ws, r, "Resultado da renda variável vs. preço médio", res, RS, cor=True)
    r = par(ws, r, "Alertas ativos agora", len(alertas), N0, f"{len(novos)} novo(s) nesta atualização")
    r += 1
    r = secao(ws, r, "ALOCAÇÃO POR CLASSE")
    classes = {"Renda Fixa": total_rf}
    for c in ("Ações", "FIIs", "Outros"):
        classes[c] = float(rv.loc[rv["Classe"] == c, "Valor"].sum()) if len(rv) else 0.0
    aloc = []
    for c, v in classes.items():
        if v == 0 and c not in cart["metas"]:
            continue
        atual = v / total if total else math.nan
        alvo = cart["metas"].get(c, math.nan)
        aloc.append({"Classe": c, "Valor (R$)": v, "Atual": atual, "Seu alvo": alvo,
                     "Desvio (p.p.)": (atual - alvo) if alvo == alvo else math.nan,
                     "Ajuste p/ voltar ao alvo (R$)": (alvo - atual) * total if alvo == alvo else math.nan})
    if aloc:
        r = tabela(ws, r, pd.DataFrame(aloc), {"Valor (R$)": RS, "Atual": PC, "Seu alvo": PC,
                   "Desvio (p.p.)": PC, "Ajuste p/ voltar ao alvo (R$)": RS},
                   cores=["Ajuste p/ voltar ao alvo (R$)"], filtro=False)
        ws.cell(r - 1, 1, "Ajuste positivo = valor a aportar na classe para voltar ao alvo que você definiu na aba Metas.").font = F_NOTA
        r += 1
    r = secao(ws, r, "MERCADO")
    for sim, nome, fmt in (("^BVSP", "Ibovespa (pontos)", N0), ("IFIX.SA", "IFIX (pontos)", N2), ("BRL=X", "Dólar comercial (R$)", '#,##0.0000')):
        y = cot.get(sim)
        if y:
            r = par(ws, r, nome, y["Preco"], fmt, f"{fpc(y['VarDia'], 2, True)} no dia" if y["VarDia"] == y["VarDia"] else None)
    for nome, (v, d) in macro.items():
        r = par(ws, r, nome, v, '#,##0.00' if "Dólar" not in nome else '#,##0.0000', f"referência {d}")
    r += 1
    if focus is not None and len(focus):
        r = secao(ws, r, "EXPECTATIVAS DE MERCADO — BOLETIM FOCUS (mediana)")
        fx = focus.pivot_table(index="Indicador", columns="Ano", values="Mediana", aggfunc="first").reset_index()
        fx.columns = [str(c) for c in fx.columns]
        r = tabela(ws, r, fx, {c: N2 for c in fx.columns[1:]}, filtro=False)
    r = secao(ws, r, "SITUAÇÃO DAS FONTES DE DADOS")
    st = pd.DataFrame(STATUS, columns=["Fonte", "Situação", "Detalhe"])
    r = tabela(ws, r, st, filtro=False, destaque=("Situação", lambda v: VERDE_F if v == "OK" else VERM_F if v == "FALHA" else AMAR_F))
    if aviso_pasta:
        ws.cell(r, 1, aviso_pasta).font = F_NEG
        r += 1
    if envio:
        ws.cell(r, 1, "Envio de alertas: " + ", ".join(f"{k}: {v}" for k, v in envio.items())).font = F_NOTA
        r += 1
    ws.cell(r + 1, 1, AVISO).font = F_NOTA

    # --------------------------------------------------------- Minha Carteira
    ws = wb.create_sheet("Minha Carteira")
    titulo(ws, "MINHA CARTEIRA", 17)
    r = 4
    r = secao(ws, r, "RENDA VARIÁVEL")
    if len(rv):
        m = pd.DataFrame({
            "Ticker": rv["Ticker"], "Classe": rv["Classe"], "Qtd": rv["Qtd"],
            "Preço médio": rv["PM"], "Cotação": rv["Preco"], "Var. dia": rv["VarDia"],
            "Valor atual (R$)": rv["Valor"], "Resultado (R$)": rv["Resultado"],
            "Resultado (%)": rv["ResultadoPct"], "Peso na carteira": rv["Peso"],
            "DY 12m": rv.get("DY"), "P/VP": rv.get("PVP"), "P/L": rv.get("PL"),
            "Score (0-100)": rv.get("Score"), "Preço justo Graham": rv.get("Graham"),
            "Preço-teto Bazin (6%)": rv.get("Bazin"),
            "Pontos de atenção (critérios objetivos)": rv["Sinais"],
            "Leitura do ranking": rv.get("Motivo")})
        r = tabela(ws, r, m.sort_values("Valor atual (R$)", ascending=False),
                   {"Qtd": N0, "Preço médio": RS, "Cotação": RS, "Var. dia": PC, "Valor atual (R$)": RS,
                    "Resultado (R$)": RS, "Resultado (%)": PC, "Peso na carteira": PC, "DY 12m": PC,
                    "P/VP": N2, "P/L": N2, "Score (0-100)": '0.0', "Preço justo Graham": RS,
                    "Preço-teto Bazin (6%)": RS},
                   cores=["Var. dia", "Resultado (R$)", "Resultado (%)"],
                   larguras={"Pontos de atenção (critérios objetivos)": 60, "Leitura do ranking": 46,
                             "Valor atual (R$)": 16, "Resultado (R$)": 15},
                   destaque=("Pontos de atenção (critérios objetivos)",
                             lambda v: None if not v or v.startswith("Sem pontos") else AMAR_F))
        ws.cell(r - 1, 1, "Graham = raiz(22,5 × lucro por ação × valor patrimonial por ação). Bazin = dividendos de 12 meses ÷ 6%. "
                          "São fórmulas clássicas de referência, não previsões de preço.").font = F_NOTA
        r += 1
    else:
        ws.cell(r, 1, "Nenhum ativo informado na aba Renda Variavel do arquivo Minha_Carteira.xlsx.").font = F_N
        r += 2
    r = secao(ws, r, "RENDA FIXA")
    if len(rf):
        r = tabela(ws, r, rf, {"Aplicado (R$)": RS, "Data aplicação": DT, "Valor bruto estimado (R$)": RS,
                               "Rendimento (R$)": RS, "Rentab. (%)": PC, "IR estimado": '0.0%',
                               "Valor líquido estimado (R$)": RS, "Vencimento": DT, "Dias p/ vencer": N0},
                   cores=["Rendimento (R$)"], filtro=False)
        ws.cell(r - 1, 1, "Valores estimados na curva do papel com CDI, Selic e IPCA oficiais; dias úteis sem feriados. "
                          "Não considera marcação a mercado nem IOF.").font = F_NOTA
    else:
        ws.cell(r, 1, "Nenhum investimento informado na aba Renda Fixa do arquivo Minha_Carteira.xlsx.").font = F_N
    ws.freeze_panes = "B4"

    # ----------------------------------------------------------------- Compras
    ws = wb.create_sheet("Compras")
    titulo(ws, "MINHAS COMPRAS E VENDAS", 12)
    r = secao(ws, 4, "HISTÓRICO DE OPERAÇÕES (da aba Compras de Minha_Carteira.xlsx)")
    lotes = cart.get("lotes") or []
    if lotes:
        precos = dict(zip(rv["Ticker"], rv["Preco"])) if len(rv) else {}
        linhas_l = []
        for x in lotes:
            cot_ = precos.get(x["Ticker"], math.nan)
            comp = x["Tipo"] == "Compra"
            bruto = x["Qtd"] * x["Preco"]
            valor_op = bruto + x["Taxas"] if comp else bruto - x["Taxas"]
            atual = x["Qtd"] * cot_ if comp and cot_ == cot_ else math.nan
            res = atual - valor_op if atual == atual else math.nan
            linhas_l.append({"Data": x["Data"], "Ticker": x["Ticker"], "Tipo": x["Tipo"], "Quantidade": x["Qtd"],
                             "Preço unitário (R$)": x["Preco"], "Taxas (R$)": x["Taxas"],
                             "Valor da operação (R$)": valor_op, "Cotação atual (R$)": cot_ if comp else math.nan,
                             "Valor atual (R$)": atual, "Resultado (R$)": res,
                             "Resultado (%)": res / valor_op if res == res and valor_op else math.nan})
        r = tabela(ws, r, pd.DataFrame(linhas_l),
                   {"Data": DT, "Quantidade": N0, "Preço unitário (R$)": RS, "Taxas (R$)": RS,
                    "Valor da operação (R$)": RS, "Cotação atual (R$)": RS, "Valor atual (R$)": RS,
                    "Resultado (R$)": RS, "Resultado (%)": PC},
                   cores=["Resultado (R$)", "Resultado (%)"], filtro=False)
        ws.cell(r - 1, 1, "Resultado = valor atual da compra menos o valor pago (com taxas). Vendas aparecem só como registro; "
                          "o preço médio da carteira segue o custo médio e não muda com vendas.").font = F_NOTA
    else:
        ws.cell(r, 1, "Nenhuma operação registrada na aba Compras de Minha_Carteira.xlsx.").font = F_N
    ws.freeze_panes = "A4"

    # ---------------------------------------------------------------- Projecao
    ws = wb.create_sheet("Projecao")
    titulo(ws, "PROJEÇÃO DA RENDA FIXA", 15)
    try:
        pf, pp, prem = projetar_rf(cart["rf"], rf, cdi_aa, focus)
    except Exception as e:  # noqa
        log.warning("Projecao: %s", e)
        pf = pp = prem = None
    if pf is None:
        ws["A4"] = "Projeção indisponível nesta atualização (sem renda fixa cadastrada ou sem CDI)."
        ws["A4"].font = F_N
    else:
        fm = {c: RS for c in pf.columns if c not in ("Investimento", "Remuneração")}
        r = secao(ws, 4, "CENÁRIO 1 - JUROS CAEM CONFORME O BOLETIM FOCUS (valor líquido estimado, após IR)")
        r = tabela(ws, r, pf, fm, larguras={"Investimento": 28, "Remuneração": 16}, filtro=False)
        r = secao(ws, r, "CENÁRIO 2 - JUROS FICAM PARADOS NO NÍVEL DE HOJE (valor líquido estimado, após IR)")
        r = tabela(ws, r, pp, fm, filtro=False)
        ganho = pf.iloc[-1]["12 meses"] - pf.iloc[-1]["Hoje (líquido, R$)"]
        g = f"{ganho:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        ws.cell(r, 1, f"Rendimento líquido estimado em 12 meses (cenário 1, sem novos aportes): R$ {g}").font = F_SEC
        r += 2
        sel = f"{prem['sel_fim']:.2f}".replace(".", ",") if prem["sel_fim"] else "n/d"
        cd = f"{prem['cdi0'] * 100:.2f}".replace(".", ",")
        notas = [
            f"Premissas: CDI de hoje {cd}% a.a. (o cenário 2 o mantém). No cenário 1 o CDI converge linearmente "
            f"até a Selic do Focus de {sel}% menos 0,10 p.p. em 31/12/{prem['alvo'].year}, e depois fica constante.",
            "Partem do valor atual estimado, acrescentam só rendimento (sem novos aportes) e descontam o IR regressivo "
            "(22,5% até 180 dias, 20% até 360, 17,5% até 720, 15% acima) sobre o rendimento total desde a aplicação.",
            "Se o título tem vencimento, o rendimento para na data de vencimento. Não considera IOF nem marcação a mercado.",
            "O Focus é uma expectativa de mercado, não garantia. Projeção estimativa para planejamento, não promessa de rentabilidade.",
        ]
        for n in notas:
            ws.cell(r, 1, n).font = F_NOTA
            r += 1
    ws.freeze_panes = "A4"

    # ----------------------------------------------------------------- Resumo
    try:
        ws = wb.create_sheet("Resumo", 1)
        titulo(ws, "RESUMO DA CARTEIRA", 8)
        linhas = []
        for i, (_, l) in enumerate(rf.iterrows()):
            apl, liq = float(l["Aplicado (R$)"]), float(l["Valor líquido estimado (R$)"])
            proj = None
            if pf is not None and i < len(pf) - 1:
                proj = float(pf.iloc[i]["12 meses"])
            linhas.append({"Classe": "Renda fixa", "Investimento": l["Investimento"], "Remuneração": l["Remuneração"],
                           "Investido (R$)": apl, "Valor atual líquido estimado (R$)": liq,
                           "Resultado (R$)": liq - apl, "Resultado (%)": (liq / apl - 1) if apl else None,
                           "Projeção em 12 meses (R$)": proj})
        for _, l in (rv.iterrows() if len(rv) else []):
            custo = l["Qtd"] * l["PM"] if l["PM"] == l["PM"] else None
            val = l["Valor"]
            linhas.append({"Classe": l["Classe"], "Investimento": l["Ticker"], "Remuneração": "preço de mercado",
                           "Investido (R$)": custo, "Valor atual líquido estimado (R$)": val,
                           "Resultado (R$)": l["Resultado"], "Resultado (%)": l["ResultadoPct"],
                           "Projeção em 12 meses (R$)": None})
        res = pd.DataFrame(linhas)
        if len(res):
            tot_v = res["Valor atual líquido estimado (R$)"].sum()
            tot_i = res["Investido (R$)"].sum()
            res.insert(len(res.columns) - 1, "% da carteira", res["Valor atual líquido estimado (R$)"] / tot_v if tot_v else None)
            tot = {c: None for c in res.columns}
            tot.update({"Classe": "TOTAL", "Investido (R$)": tot_i, "Valor atual líquido estimado (R$)": tot_v,
                        "Resultado (R$)": tot_v - tot_i, "Resultado (%)": (tot_v / tot_i - 1) if tot_i else None,
                        "% da carteira": 1.0, "Projeção em 12 meses (R$)":
                        (pf.iloc[-1]["12 meses"] if pf is not None and not len(rv) else None)})
            res = pd.concat([res, pd.DataFrame([tot])], ignore_index=True)
            r = tabela(ws, 4, res, {"Investido (R$)": RS, "Valor atual líquido estimado (R$)": RS, "Resultado (R$)": RS,
                                    "Resultado (%)": PC, "% da carteira": PC, "Projeção em 12 meses (R$)": RS},
                       cores=["Resultado (R$)", "Resultado (%)"], filtro=False,
                       larguras={"Classe": 14, "Investimento": 30, "Remuneração": 18,
                                 "Valor atual líquido estimado (R$)": 22, "Projeção em 12 meses (R$)": 20})
            for j in range(1, len(res.columns) + 1):
                ws.cell(4 + len(res), j).font = F_SEC
            ws.cell(r - 1, 1, "Renda fixa: valor estimado na curva do papel, líquido de IR, com CDI/IPCA oficiais; atualiza sozinho a cada rodada. "
                              "Renda variável: cotação de mercado. A projeção de 12 meses vem da aba Projecao (cenário Focus, só renda fixa, sem novos aportes).").font = F_NOTA
        else:
            ws["A4"] = "Nenhum investimento informado em Minha_Carteira.xlsx."
            ws["A4"].font = F_N
        ws.freeze_panes = "A5"
    except Exception as e:  # noqa
        log.warning("Resumo: %s", e)

    # ---------------------------------------------------------------- Alertas
    ws = wb.create_sheet("Alertas")
    titulo(ws, "ALERTAS DA CARTEIRA", 3)
    al = pd.DataFrame([{"Nível": n, "Alerta": t, "Situação": ("enviado agora" if enviou else "novo (envio desligado ou com erro)")
                        if k in chaves_novas else "já enviado"}
                       for k, n, t in alertas], columns=["Nível", "Alerta", "Situação"])
    if len(al):
        tabela(ws, 4, al, larguras={"Nível": 10, "Alerta": 110, "Situação": 34},
               destaque=("Nível", lambda v: VERM_F if v == "ALTO" else AMAR_F))
    else:
        ws["A4"] = "Nenhum alerta pelos limites configurados."
        ws["A4"].font = F_N

    # --------------------------------------------------------------- Notícias
    ws = wb.create_sheet("Noticias")
    titulo(ws, "NOTÍCIAS DOS MEUS ATIVOS (últimos 7 dias)", 5)
    if len(noticias):
        tabela(ws, 4, noticias, {"Data": 'DD/MM/YYYY HH:MM'}, link="Link",
               larguras={"Ticker": 10, "Data": 17, "Título": 95, "Fonte": 24, "Relevante": 11},
               destaque=("Relevante", lambda v: VERM_F if v == "ALTA" else AMAR_F if v == "MÉDIA" else None))
    else:
        ws["A4"] = "Sem notícias encontradas."
    ws.freeze_panes = "A5"

    # ------------------------------------------------------------- Rankings
    def aba_ranking(nome, tit, df, cols, fmts, grupos, nota):
        w = wb.create_sheet(nome)
        titulo(w, tit, len(cols) + 2)
        w["A3"] = nota
        w["A3"].font = F_NOTA
        d = df.dropna(subset=["Score"]).sort_values("Score", ascending=False).head(30).copy()
        if not len(d):
            w["A5"] = "Dados indisponíveis nesta atualização."
            return
        d.insert(0, "Posição", range(1, len(d) + 1))
        d["Na carteira?"] = d["Ticker"].map(lambda t: "SIM" if t in meus else "")
        d["Por que aparece"] = d.apply(lambda l: motivo(l, grupos), axis=1)
        sel = d[["Posição", "Ticker"] + list(cols.keys()) + ["Na carteira?", "Por que aparece"]].rename(columns=cols)
        tabela(w, 5, sel, {cols.get(k, k): v for k, v in fmts.items()}, cores=[cols.get("VarDia")],
               larguras={"Por que aparece": 52, "Posição": 9},
               destaque=("Na carteira?", lambda v: VERDE_F if v == "SIM" else None))
        w.freeze_panes = "C6"

    nota_rank = ("Ranking por critérios objetivos, com pesos ajustáveis no config.ini. Serve para filtrar "
                 "candidatos a estudo; não é recomendação de compra.")
    aba_ranking("Ranking Acoes", "RANKING DE AÇÕES — 30 MELHORES SCORES", acoes,
                {"Preco": "Cotação", "VarDia": "Var. dia", "Score": "Score (0-100)", "PL": "P/L", "PVP": "P/VP",
                 "DY": "DY 12m", "ROE": "ROE", "DivLiq_Patrim": "Dív.líq/Patrim.", "Cresc5a": "Cresc. receita 5a",
                 "MargemGraham": "Margem vs. Graham", "Liquidez": "Liquidez diária (R$)"},
                {"Preco": RS, "VarDia": PC, "Score": '0.0', "PL": N2, "PVP": N2, "DY": PC, "ROE": PC,
                 "DivLiq_Patrim": N2, "Cresc5a": PC, "MargemGraham": PC, "Liquidez": N0},
                ["Valor", "Renda", "Qualidade", "Crescimento"], nota_rank) if len(acoes) else None
    aba_ranking("Ranking FIIs", "RANKING DE FUNDOS IMOBILIÁRIOS — 30 MELHORES SCORES", fiis,
                {"Segmento": "Segmento", "Preco": "Cotação", "VarDia": "Var. dia", "Score": "Score (0-100)",
                 "DY": "DY 12m", "PVP": "P/VP", "Vacancia": "Vacância", "QtdImoveis": "Imóveis",
                 "Liquidez": "Liquidez diária (R$)"},
                {"Preco": RS, "VarDia": PC, "Score": '0.0', "DY": PC, "PVP": N2, "Vacancia": PC,
                 "QtdImoveis": N0, "Liquidez": N0},
                ["Renda", "Preço", "Liquidez", "Vacância"], nota_rank) if len(fiis) else None

    # ------------------------------------------------------------ Renda Fixa
    ws = wb.create_sheet("Renda Fixa")
    titulo(ws, "RENDA FIXA — REFERÊNCIAS DE MERCADO", 7)
    r = 4
    if cdi_frac:
        r = secao(ws, r, f"QUANTO RENDE HOJE (CDI de {fpc(cdi_frac, 2)} a.a.)")
        ref = []
        for nome, bruto, isento in (("CDB 100% do CDI", cdi_frac, False), ("CDB 110% do CDI", cdi_frac * 1.10, False),
                                    ("LCI/LCA 90% do CDI (isento)", cdi_frac * 0.90, True),
                                    ("Tesouro Selic (aprox. Selic)", (selic or cdi_aa) / 100, False),
                                    ("Poupança (aprox., sem TR)", 0.0617 if (selic or 0) > 8.5 else 0.7 * (selic or 0) / 100, True)):
            ref.append({"Aplicação": nome, "Taxa bruta a.a.": bruto,
                        "Líquido até 6 meses (IR 22,5%)": bruto if isento else bruto * 0.775,
                        "Líquido 1 a 2 anos (IR 17,5%)": bruto if isento else bruto * 0.825,
                        "Líquido acima de 2 anos (IR 15%)": bruto if isento else bruto * 0.85})
        r = tabela(ws, r, pd.DataFrame(ref), {c: PC for c in ref[0] if c != "Aplicação"},
                   larguras={"Aplicação": 34}, filtro=False)
        ipca_prox = None
        if focus is not None and len(focus):
            f2 = focus[(focus["Indicador"] == "IPCA") & (focus["Ano"] == str(HOJE.year + 1))]
            ipca_prox = float(f2["Mediana"].iloc[0]) / 100 if len(f2) else None
        if ipca_prox:
            r = par(ws, r, "Juro real esperado (CDI ÷ IPCA Focus do próximo ano)", (1 + cdi_frac) / (1 + ipca_prox) - 1, PC,
                    f"IPCA esperado: {fpc(ipca_prox, 2)}")
            r += 1
    r = secao(ws, r, "TESOURO DIRETO — TAXAS DE REFERÊNCIA")
    if tesouro is not None and len(tesouro):
        t = tesouro.copy()
        t["Prazo (anos)"] = t["Vencimento"].map(lambda d: (d - HOJE).days / 365.25)
        t["Taxa compra (% a.a.)"] = t["Taxa compra (% a.a.)"] / 100
        t = t[["Título", "Vencimento", "Prazo (anos)", "Taxa compra (% a.a.)", "PU compra (R$)", "Data base"]]
        tabela(ws, r, t, {"Vencimento": DT, "Prazo (anos)": '0.0', "Taxa compra (% a.a.)": PC,
                          "PU compra (R$)": RS, "Data base": DT}, larguras={"Título": 40})
    else:
        ws.cell(r, 1, "Taxas do Tesouro indisponíveis nesta atualização.").font = F_N
    ws.column_dimensions["A"].width = 46
    for col in "BCDEF":
        ws.column_dimensions[col].width = 22

    # ------------------------------------------------------------- Mercado
    def aba_mercado(nome, tit, df, cols, fmts):
        w = wb.create_sheet(nome)
        titulo(w, tit, len(cols) + 1)
        if not len(df):
            w["A4"] = "Dados indisponíveis nesta atualização."
            return
        d = df.sort_values("Liquidez", ascending=False)[["Ticker"] + list(cols.keys())].rename(columns=cols)
        tabela(w, 4, d, {cols[k]: v for k, v in fmts.items()}, cores=[cols["VarDia"]])
        w.freeze_panes = "B5"

    aba_mercado("Acoes", f"AÇÕES — {len(acoes)} MAIS LÍQUIDAS DA B3", acoes,
                {"Preco": "Cotação", "VarDia": "Var. dia", "Min52": "Mín. 52 sem.", "Max52": "Máx. 52 sem.",
                 "PL": "P/L", "PVP": "P/VP", "DY": "DY 12m", "EV_EBIT": "EV/EBIT", "MrgLiq": "Margem líq.",
                 "ROE": "ROE", "ROIC": "ROIC", "DivLiq_Patrim": "Dív.líq/Patrim.", "Cresc5a": "Cresc. receita 5a",
                 "Graham": "Preço justo Graham", "Bazin": "Preço-teto Bazin", "Score": "Score (0-100)",
                 "Liquidez": "Liquidez diária (R$)"},
                {"Preco": RS, "VarDia": PC, "Min52": RS, "Max52": RS, "PL": N2, "PVP": N2, "DY": PC,
                 "EV_EBIT": N2, "MrgLiq": PC, "ROE": PC, "ROIC": PC, "DivLiq_Patrim": N2, "Cresc5a": PC,
                 "Graham": RS, "Bazin": RS, "Score": '0.0', "Liquidez": N0})
    aba_mercado("FIIs", f"FUNDOS IMOBILIÁRIOS — {len(fiis)} MAIS LÍQUIDOS DA B3", fiis,
                {"Segmento": "Segmento", "Preco": "Cotação", "VarDia": "Var. dia", "Min52": "Mín. 52 sem.",
                 "Max52": "Máx. 52 sem.", "DY": "DY 12m", "FFOYield": "FFO yield", "PVP": "P/VP",
                 "Vacancia": "Vacância", "QtdImoveis": "Imóveis", "CapRate": "Cap rate",
                 "ValorMercado": "Valor de mercado (R$)", "Score": "Score (0-100)", "Liquidez": "Liquidez diária (R$)"},
                {"Preco": RS, "VarDia": PC, "Min52": RS, "Max52": RS, "DY": PC, "FFOYield": PC, "PVP": N2,
                 "Vacancia": PC, "QtdImoveis": N0, "CapRate": PC, "ValorMercado": N0, "Score": '0.0',
                 "Liquidez": N0})

    final = salvar(wb, arq_painel)
    try:
        import plano
        log.info("Plano de acao gravado em %s", plano.gerar(pasta, final))
    except Exception as e:  # noqa
        log.exception("Falha ao gerar o Plano de Acao: %s", e)
    log.info("Painel gravado em %s | %d alertas (%d novos) | envio: %s", final, len(alertas), len(novos), envio)


if __name__ == "__main__":
    try:
        main()
    except Exception:  # noqa
        logging.getLogger("robo").exception("Erro fatal")
        raise
