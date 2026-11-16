#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PLANO DE ACAO
Le o Painel_Investimentos.xlsx e gera o Plano_de_Acao.xlsx: sinais de compra, venda e
manutencao calculados pelas regras da aba "Parametros" (que voce edita).

Os sinais sao o resultado mecanico dos SEUS criterios sobre dados publicos.
Nao constituem recomendacao de investimento.

Uso isolado:  python plano.py "<pasta com o Painel>"
"""
import datetime as dt
import logging
import math
import os
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation

AGORA = dt.datetime.now()
from pathlib import Path  # noqa: E402
BASE = Path(__file__).resolve().parent
log = logging.getLogger("robo")
ARQ_PLANO = "Plano_de_Acao.xlsx"

# Alvos por perfil (Renda Fixa, Acoes, FIIs), em %. Podem ser editados na aba Parametros.
ALVOS_PERFIL = {"Conservador": (80.5, 8.0, 11.5), "Moderado": (65.5, 14.0, 20.5),
                "Arrojado": (46.0, 22.0, 32.0), "Personalizado": (80.5, 8.0, 11.5)}

# (chave, rotulo, valor padrao, explicacao)
PARAMETROS = [
    ("sec1", "ALOCAÇÃO-ALVO E APORTE", None, None),
    ("perfil", "Perfil de risco", "Conservador",
     "Escolha na lista: Conservador, Moderado, Arrojado ou Personalizado. Os três alvos abaixo mudam sozinhos, conforme a tabela no fim desta aba."),
    ("alvo_rf", "Alvo em Renda Fixa (%)", None, "Calculado pelo perfil escolhido (não digite aqui; para mudar, edite a tabela de alvos por perfil abaixo)."),
    ("alvo_acoes", "Alvo em Ações (%)", None, "Calculado pelo perfil escolhido."),
    ("alvo_fiis", "Alvo em FIIs (%)", None, "Calculado pelo perfil escolhido."),
    ("aporte", "Aporte disponível agora (R$)", 0, "Dinheiro novo a investir. É a mesma célula do Dashboard: digite lá (célula amarela) ou aqui por cima; a aba Plano de Acao se recalcula na hora, no Excel."),
    ("max_ativos", "Nº máximo de ativos por classe no plano de compras", 5, "O valor de cada classe é dividido igualmente entre os melhores candidatos."),
    ("conc_max", "Concentração máxima por ativo (% da carteira)", 10, "Acima disso o sinal vira REDUZIR."),
    ("div_cresc", "Crescimento anual dos dividendos nas projeções do Dashboard (%)", 0,
     "Usado só na soma dos dividendos de 2 a 10 anos. 0 = dividendos constantes. Estimativa, não promessa."),
    ("sec2", "AÇÕES — CRITÉRIOS DE COMPRA (todos precisam ser atendidos)", None, None),
    ("a_liq", "Liquidez diária mínima (R$)", 5000000, None),
    ("a_pl_max", "P/L máximo", 15, "Preço ÷ lucro por ação."),
    ("a_pvp_max", "P/VP máximo", 3, None),
    ("a_roe_min", "ROE mínimo (%)", 12, None),
    ("a_div_max", "Dívida líquida ÷ patrimônio máxima", 1.5, None),
    ("a_mrg_min", "Margem líquida mínima (%)", 10, "A margem líquida precisa ser MAIOR que este valor."),
    ("a_dle_max", "Dívida líquida ÷ EBITDA máxima", 3, "Precisa ser MENOR que este valor. Se o Yahoo Finance não tiver o dado da empresa, este critério é ignorado nela."),
    ("a_dy_min", "Dividend yield mínimo (%)", 4, None),
    ("a_dy_max", "Dividend yield máximo (%)", 18, "Filtro contra armadilhas: DY muito alto costuma vir de dividendo extraordinário, que não se repete."),
    ("a_pl_min", "P/L mínimo", 3, "P/L muito baixo costuma vir de lucro não recorrente (venda de ativo, efeito contábil)."),
    ("a_premio", "Prêmio mínimo do lucro sobre o juro real (p.p.)", 3,
     "Lucro ÷ preço precisa superar a taxa do Tesouro IPCA+ mais este prêmio. É o que liga o critério à expectativa de mercado."),
    ("a_score_min", "Score mínimo (0-100)", 60, None),
    ("sec3", "FIIs — CRITÉRIOS DE COMPRA (todos precisam ser atendidos)", None, None),
    ("f_liq", "Liquidez diária mínima (R$)", 1000000, None),
    ("f_premio", "Prêmio mínimo do DY sobre o Tesouro IPCA+ (p.p.)", 2,
     "DY precisa superar a taxa do Tesouro IPCA+ mais este prêmio."),
    ("f_dy_max", "Dividend yield máximo (%)", 16, "Filtro contra armadilhas: DY muito alto costuma indicar problema ou evento não recorrente."),
    ("f_pvp_min", "P/VP mínimo", 0.80, "Cota muito abaixo do patrimônio costuma indicar desconfiança do mercado."),
    ("f_pvp_max", "P/VP máximo", 1.05, None),
    ("f_vac_max", "Vacância máxima (%)", 12, None),
    ("f_score_min", "Score mínimo (0-100)", 55, None),
    ("sec4", "GATILHOS DE VENDA (basta um)", None, None),
    ("v_score", "Score abaixo de", 35, "Vale para ações e FIIs."),
    ("v_pl_max", "Ações: P/L acima de (ou lucro negativo)", 30, None),
    ("v_div_max", "Ações: dívida líquida ÷ patrimônio acima de", 3, None),
    ("v_pvp_fii", "FIIs: P/VP acima de", 1.15, None),
    ("v_vac_fii", "FIIs: vacância acima de (%)", 20, None),
    ("sec5", "METAS E LIMITES DE RISCO (Dashboard)", None, None),
    ("meta_pat", "Meta de patrimônio (R$)", 100000, "Alimenta a seção Meta e progresso do Dashboard."),
    ("meta_renda", "Meta de renda mensal (R$)", 10000, "Renda = rendimento da renda fixa + proventos de ações e FIIs."),
    ("meta_prazo", "Prazo para a meta (anos)", 1, None),
    ("meta_ret", "Retorno líquido esperado da carteira (% ao ano)", 10, "Usado só nas contas da meta (aporte mensal necessário e tempo até a meta). É uma hipótese sua, não uma promessa."),
    ("r_fgc", "Limite de cobertura do FGC por instituição (R$)", 250000, "O FGC cobre até este valor por CPF e por instituição (e R$ 1 milhão a cada 4 anos no total)."),
    ("r_inst", "Alerta de concentração: % do patrimônio em uma instituição acima de", 50, None),
    ("r_venc", "Alerta de vencimento: títulos que vencem em até (dias)", 60, None),
]

DASH_PERFIL = "B4"      # seletor de perfil no Dashboard
DASH_APORTE = "B24"      # aporte disponivel (mesma escolha da aba Parametros)
DASH_H_RF = "L11"       # horizonte do card de projecao da renda fixa
DASH_H_DV = "L15"       # horizonte do card de dividendos
HORIZ_ROTULOS = ["12 meses"] + [f"{a} anos" for a in range(2, 11)]

AVISO = ("Sinais calculados mecanicamente pelos critérios da aba Parametros sobre dados públicos. "
         "Não constituem recomendação de investimento; a decisão é sua. Para orientação personalizada, "
         "procure um profissional credenciado pela CVM.")

# ------------------------------------------------------------------ estilo
ARIAL = "Arial"
F_N = Font(name=ARIAL, size=10)
F_B = Font(name=ARIAL, size=10, bold=True)
F_CAB = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
F_TIT = Font(name=ARIAL, size=15, bold=True, color="FFFFFF")
F_SEC = Font(name=ARIAL, size=11, bold=True, color="1F3864")
F_NOTA = Font(name=ARIAL, size=9, italic=True, color="595959")
F_ENT = Font(name=ARIAL, size=10, color="0000FF")
AZUL = PatternFill("solid", fgColor="1F3864")
VERDE = PatternFill("solid", fgColor="C6E0B4")
VERM = PatternFill("solid", fgColor="F8CBAD")
AMAR = PatternFill("solid", fgColor="FFE699")
CINZA = PatternFill("solid", fgColor="EDEDED")
ENTRADA = PatternFill("solid", fgColor="FFF2CC")
BORDA = Border(bottom=Side(style="thin", color="D9D9D9"))
RS, PC, N2, N0 = 'R$ #,##0.00', '0.00%', '#,##0.00', '#,##0'
COR_SINAL = {"COMPRAR": VERDE, "COMPRAR MAIS": VERDE, "VENDER": VERM, "REDUZIR": AMAR, "MANTER": CINZA}


def fbr(v, dec=2):
    return f"{v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fpc(v, dec=1):
    return fbr(v * 100, dec) + "%"


# colunas das abas de candidatos (usadas na planilha e na ficha do Dashboard)
FICHA_FMT = {"Cotação": "rs", "Score (0-100)": "n1", "P/L": "n2", "Lucro ÷ preço": "pc", "P/VP": "n2", "DY 12m": "pc", "Rendimento 12m (R$/cota)": "rs",
             "ROE": "pc", "Margem líq.": "pc", "Dív.líq/Patrim.": "n2", "Dív.líq/EBITDA": "n2", "EBITDA (R$ mi)": "n0", "Dív. líquida (R$ mi)": "n0", "Preço justo Graham": "rs", "Margem vs. Graham": "pc",
             "Preço-teto Bazin": "rs", "Liquidez diária (R$)": "n0", "Prêmio sobre IPCA+ (p.p.)": "n1",
             "Vacância": "pc", "Imóveis": "n0"}
COLS_AC = ["Cotação", "Score (0-100)", "P/L", "Lucro ÷ preço", "P/VP", "DY 12m", "Rendimento 12m (R$/cota)", "ROE", "Margem líq.", "Dív.líq/Patrim.", "Dív.líq/EBITDA", "EBITDA (R$ mi)", "Dív. líquida (R$ mi)",
           "Preço justo Graham", "Margem vs. Graham", "Preço-teto Bazin", "Liquidez diária (R$)"]


def cols_fi(cf):
    return (["Segmento", "Cotação", "Score (0-100)", "DY 12m", "Rendimento 12m (R$/cota)"]
            + (["Prêmio sobre IPCA+ (p.p.)"] if "Prêmio sobre IPCA+ (p.p.)" in cf.columns else [])
            + ["P/VP", "Vacância", "Imóveis", "Liquidez diária (R$)"])


def _fmt_ficha(v, k):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "n/d"
    if k is None:
        return str(v)
    try:
        v = float(v)
    except (TypeError, ValueError):
        return str(v)
    return {"rs": "R$ " + fbr(v), "pc": fbr(v * 100) + "%", "n2": fbr(v), "n1": fbr(v, 1), "n0": fbr(v, 0)}[k]


def ficha_linha(l, cols):
    """Todos os campos da linha da aba de candidatos, ja formatados: [[rotulo, texto], ...]."""
    campos = ["Sinal", "Ticker", "Critérios atendidos"] + cols + ["Não atende em"]
    out = []
    for c in campos:
        v = l.get(c) if hasattr(l, "get") else None
        out.append([c, _fmt_ficha(v, FICHA_FMT.get(c)) if v is not None else "n/d"])
    return out


def ok(v):
    return v is not None and isinstance(v, (int, float)) and not (isinstance(v, float) and math.isnan(v))


# ------------------------------------------------------- leitura do painel
def _val(cel):
    """Valor da celula; se for numero formatado como %, devolve em pontos (5% -> 5)."""
    v = cel.value
    if isinstance(v, (int, float)) and "%" in (cel.number_format or ""):
        return v * 100
    return v


def pares(ws):
    out = {}
    for lin in ws.iter_rows(max_col=2):
        a, b = lin[0].value, _val(lin[1]) if len(lin) > 1 else None
        if isinstance(a, str) and isinstance(b, (int, float)):
            out[a.strip()] = b
    return out


def tabela(ws, chave, pct_como_pontos=False):
    """Tabela cujo cabecalho comeca com `chave` na coluna A, ate a primeira linha vazia."""
    linhas = list(ws.iter_rows())
    ini = next((i for i, l in enumerate(linhas) if l[0].value == chave), None)
    if ini is None:
        return pd.DataFrame()
    cab = []
    for c in linhas[ini]:
        if c.value is None:
            break
        cab.append(str(c.value))
    dados = []
    for l in linhas[ini + 1:]:
        if l[0].value is None or (len(cab) > 1 and all(c.value is None for c in l[1:len(cab)])):
            break  # linha vazia ou nota de rodape
        dados.append([(_val(c) if pct_como_pontos else c.value) for c in l[:len(cab)]])
    return pd.DataFrame(dados, columns=cab)


def ler_painel(caminho):
    wb = load_workbook(caminho, data_only=True)
    p = {"macro": pares(wb["Painel"]),
         "aloc": tabela(wb["Painel"], "Classe"),
         "focus": tabela(wb["Painel"], "Indicador", pct_como_pontos=True),
         "rv": tabela(wb["Minha Carteira"], "Ticker"),
         "rf": tabela(wb["Minha Carteira"], "Investimento"),
         "acoes": tabela(wb["Acoes"], "Ticker") if "Acoes" in wb.sheetnames else pd.DataFrame(),
         "fiis": tabela(wb["FIIs"], "Ticker") if "FIIs" in wb.sheetnames else pd.DataFrame(),
         "tesouro": tabela(wb["Renda Fixa"], "Título") if "Renda Fixa" in wb.sheetnames else pd.DataFrame(),
         "proj": tabela(wb["Projecao"], "Investimento") if "Projecao" in wb.sheetnames else pd.DataFrame(),
         "compras": tabela(wb["Compras"], "Data") if "Compras" in wb.sheetnames else pd.DataFrame(),
         "rf_ref": tabela(wb["Renda Fixa"], "Aplicação") if "Renda Fixa" in wb.sheetnames else pd.DataFrame(),
         "atualizado": wb["Painel"]["A2"].value or ""}
    for chave, aba_, col_ in (("alertas", "Alertas", "Nível"), ("resumo", "Resumo", "Classe"), ("proventos", "Proventos", "Ticker")):
        try:
            p[chave] = tabela(wb[aba_], col_) if aba_ in wb.sheetnames else pd.DataFrame()
        except Exception:  # noqa
            p[chave] = pd.DataFrame()
    wb.close()
    return p


def ler_parametros(caminho):
    """Le os valores que o usuario editou na aba Parametros do plano anterior."""
    par = {k: v for k, _, v, _ in PARAMETROS if not k.startswith("sec")}
    par["tabela_alvos"] = {k: tuple(v) for k, v in ALVOS_PERFIL.items()}
    legado = {}
    if not Path(caminho).exists():
        return par
    try:
        wb = load_workbook(caminho, data_only=True)
        if "Parametros" in wb.sheetnames:
            rot = {r: k for k, r, _, _ in PARAMETROS}
            secao = ""
            for a, b in wb["Parametros"].iter_rows(max_col=2):
                nome = str(a.value).strip() if a.value else ""
                # rotulos repetidos (liquidez, score) sao distinguidos pela secao
                if nome in [r for k, r, _, _ in PARAMETROS if k.startswith("sec")]:
                    secao = nome
                    continue
                for k, r, _, _ in PARAMETROS:
                    if r == nome and _secao_de(k) == secao:
                        v = b.value
                        if k == "perfil":
                            if v not in (None, ""):
                                par[k] = str(v).strip()
                            continue
                        if k.startswith("alvo"):
                            legado[k] = v
                            continue
                        if v is None or v == "":
                            par[k] = None if k.startswith("alvo") else par[k]
                        elif isinstance(v, (int, float)):
                            par[k] = v * 100 if "%" in (b.number_format or "") and "(%)" in r else v
                        else:
                            try:
                                par[k] = float(str(v).replace(".", "").replace(",", "."))
                            except ValueError:
                                pass
            achou = False
            for lin in wb["Parametros"].iter_rows(max_col=4):
                nome = str(lin[0].value).strip() if lin[0].value else ""
                if nome in ALVOS_PERFIL and all(isinstance(c.value, (int, float)) for c in lin[1:4]):
                    par["tabela_alvos"][nome] = tuple(
                        c.value * 100 if "%" in (c.number_format or "") else float(c.value) for c in lin[1:4])
                    achou = True
            if not achou:  # arquivo antigo: os alvos digitados viram o perfil Personalizado
                try:
                    vals = [float(legado[k]) for k in ("alvo_rf", "alvo_acoes", "alvo_fiis")]
                    if abs(sum(vals) - 100) < 0.5:
                        par["tabela_alvos"]["Personalizado"] = tuple(vals)
                except (KeyError, TypeError, ValueError):
                    pass
        # perfil e aporte: a celula do Dashboard e a da aba Parametros sao a mesma escolha.
        # Se o usuario digitou direto na Parametros (por cima da formula), vale o que ele digitou.
        try:
            wf = load_workbook(caminho)
            digitado = {}
            if "Parametros" in wf.sheetnames:
                for lin in wf["Parametros"].iter_rows(max_col=2):
                    rot_ = str(lin[0].value).strip()
                    v = lin[1].value
                    if rot_ == "Perfil de risco" and isinstance(v, str) and not v.startswith("=") and v.strip():
                        digitado["perfil"] = v.strip()
                    if rot_ == "Aporte disponível agora (R$)" and isinstance(v, (int, float)) and not isinstance(v, bool):
                        digitado["aporte"] = float(v)
            wf.close()
            da = {}
            if "Dashboard" in wb.sheetnames:
                v = wb["Dashboard"][DASH_PERFIL].value
                if isinstance(v, str) and v.strip():
                    da["perfil"] = v.strip()
                v = wb["Dashboard"][DASH_APORTE].value
                if isinstance(v, (int, float)) and not isinstance(v, bool) and v >= 0:
                    da["aporte"] = float(v)
            for k_ in ("perfil", "aporte"):
                escolha = digitado.get(k_, da.get(k_))
                if escolha is not None:
                    par[k_] = escolha
        except Exception:  # noqa
            pass
        wb.close()
    except Exception:  # noqa
        pass
    return par


def ler_dash(caminho):
    """Horizontes escolhidos nos seletores do Dashboard (para nao voltarem ao padrao a cada atualizacao)."""
    out = {}
    try:
        if Path(caminho).exists():
            wb = load_workbook(caminho, data_only=True)
            if "Dashboard" in wb.sheetnames:
                ws = wb["Dashboard"]
                for k, cel in (("rf_card", DASH_H_RF), ("dv_card", DASH_H_DV)):
                    v = ws[cel].value
                    if v in HORIZ_ROTULOS:
                        out[k] = v
                for lin in ws.iter_rows(max_col=2):
                    a = lin[0].value
                    for k, rot in (("rf_tab", "Horizonte (renda fixa)"), ("ac_tab", "Horizonte (ações)"),
                                   ("fi_tab", "Horizonte (FIIs)")):
                        if a == rot and lin[1].value in HORIZ_ROTULOS:
                            out[k] = lin[1].value
            wb.close()
    except Exception:  # noqa
        pass
    return out


def _secao_de(chave):
    atual = ""
    for k, r, _, _ in PARAMETROS:
        if k.startswith("sec"):
            atual = r
        if k == chave:
            return atual
    return ""


def mapa_param():
    """Linha de cada parametro na aba Parametros (mesma contagem usada ao escrever a aba)."""
    r, m = 4, {}
    for k, rot, _, exp in PARAMETROS:
        if k.startswith("sec"):
            r += 1
        m[k] = r
        r += 1
    return m


PERFIS = {"Conservador": 15, "Moderado": 30, "Arrojado": 50}


def tabela_pos():
    """(linha do cabecalho, primeira linha de dados) da tabela de alvos por perfil na aba Parametros."""
    last = max(mapa_param().values())
    return last + 3, last + 4


def alvo_automatico(selic, selic_fut, prem_a, prem_f):
    """Alvo por perfil = base do perfil + ajuste pelo premio de mercado + ajuste pela tendencia da Selic."""
    prem = [x for x in (prem_a, prem_f) if x is not None]
    media = sum(prem) / len(prem) * 100 if prem else None
    t_prem = max(-5.0, min(5.0, (media - 4.0) * 1.5)) if media is not None else 0.0
    t_juros, juros_txt = 0.0, "sem dado do Focus"
    if selic is not None and selic_fut is not None:
        if selic_fut <= selic - 0.5:
            t_juros, juros_txt = 3.0, "Focus projeta queda da Selic"
        elif selic_fut >= selic + 0.5:
            t_juros, juros_txt = -5.0, "Focus projeta alta da Selic"
        else:
            juros_txt = "Focus projeta Selic estável"
    fa = 0.5
    if prem_a is not None and prem_f is not None and prem_a > 0 and prem_f > 0:
        fa = max(0.35, min(0.65, prem_a / (prem_a + prem_f)))
    tab = {}
    for nome, base in PERFIS.items():
        rv = max(max(base - 10, 0), min(base + 10, base + t_prem + t_juros))
        ac = round(rv * fa * 2) / 2
        fi = round((rv - rv * fa) * 2) / 2
        tab[nome] = {"Renda Fixa": 100 - ac - fi, "Ações": ac, "FIIs": fi}
    return tab, t_prem, t_juros, juros_txt, media, fa


# ------------------------------------------------------------------ regras
def avaliar_acao(l, par, juro_real):
    pl, ey = l.get("P/L"), None
    if ok(pl) and pl > 0:
        ey = 1 / pl
    c = [
        ("Liquidez", ok(l.get("Liquidez diária (R$)")) and l["Liquidez diária (R$)"] >= par["a_liq"]),
        (f"P/L entre {par['a_pl_min']:g} e {par['a_pl_max']:g}", ok(pl) and par["a_pl_min"] <= pl <= par["a_pl_max"]),
        (f"P/VP até {par['a_pvp_max']:g}", ok(l.get("P/VP")) and 0 < l["P/VP"] <= par["a_pvp_max"]),
        (f"ROE mín. {par['a_roe_min']:g}%", ok(l.get("ROE")) and l["ROE"] * 100 >= par["a_roe_min"]),
        (f"Margem líq. > {par['a_mrg_min']:g}%", ok(l.get("Margem líq.")) and l["Margem líq."] * 100 > par["a_mrg_min"]),
        (f"Dív.líq/EBITDA < {par['a_dle_max']:g}", (not ok(l.get("Dív.líq/EBITDA"))) or l["Dív.líq/EBITDA"] < par["a_dle_max"]),
        (f"Dívida até {par['a_div_max']:g}x", (not ok(l.get("Dív.líq/Patrim."))) or l["Dív.líq/Patrim."] <= par["a_div_max"]),
        (f"DY entre {par['a_dy_min']:g}% e {par['a_dy_max']:g}%",
         ok(l.get("DY 12m")) and par["a_dy_min"] <= l["DY 12m"] * 100 <= par["a_dy_max"]),
        ("Prêmio sobre juro real", ey is not None and juro_real is not None and ey * 100 >= juro_real * 100 + par["a_premio"]),
        (f"Score mín. {par['a_score_min']:g}", ok(l.get("Score (0-100)")) and l["Score (0-100)"] >= par["a_score_min"]),
    ]
    return c, ey


def avaliar_fii(l, par, ipca_mais):
    dy, pvp, vac = l.get("DY 12m"), l.get("P/VP"), l.get("Vacância")
    c = [
        ("Liquidez", ok(l.get("Liquidez diária (R$)")) and l["Liquidez diária (R$)"] >= par["f_liq"]),
        ("Prêmio sobre Tesouro IPCA+", ok(dy) and ipca_mais is not None and dy * 100 >= ipca_mais * 100 + par["f_premio"]),
        (f"DY até {par['f_dy_max']:g}%", ok(dy) and dy * 100 <= par["f_dy_max"]),
        (f"P/VP entre {par['f_pvp_min']:g} e {par['f_pvp_max']:g}", ok(pvp) and par["f_pvp_min"] <= pvp <= par["f_pvp_max"]),
        (f"Vacância até {par['f_vac_max']:g}%", (not ok(vac)) or vac * 100 <= par["f_vac_max"]),
        (f"Score mín. {par['f_score_min']:g}", ok(l.get("Score (0-100)")) and l["Score (0-100)"] >= par["f_score_min"]),
    ]
    return c


def gatilhos_venda(l, classe, par):
    g = []
    sc = l.get("Score (0-100)")
    if ok(sc) and sc < par["v_score"]:
        g.append(f"score {sc:.0f} abaixo de {par['v_score']:g}")
    if classe == "Ações":
        pl = l.get("P/L")
        if ok(pl) and pl <= 0:
            g.append("empresa com prejuízo em 12 meses")
        elif ok(pl) and pl > par["v_pl_max"]:
            g.append(f"P/L de {fbr(pl, 1)} acima de {par['v_pl_max']:g}")
        d = l.get("Dív.líq/Patrim.")
        if ok(d) and d > par["v_div_max"]:
            g.append(f"dívida de {fbr(d, 1)}x o patrimônio")
    else:
        pvp, vac = l.get("P/VP"), l.get("Vacância")
        if ok(pvp) and pvp > par["v_pvp_fii"]:
            g.append(f"P/VP de {fbr(pvp)} acima de {par['v_pvp_fii']:g}")
        if ok(vac) and vac * 100 > par["v_vac_fii"]:
            g.append(f"vacância de {fpc(vac, 0)}")
    return g


def candidatos(df, avaliar, *args):
    if df.empty:
        return df
    linhas = []
    for _, l in df.iterrows():
        res = avaliar(l, *args)
        crit = res[0] if isinstance(res, tuple) else res
        falhas = [n for n, v in crit if not v]
        d = l.to_dict()
        d["Critérios atendidos"] = f"{len(crit) - len(falhas)} de {len(crit)}"
        d["_n"] = len(crit) - len(falhas)
        d["_todos"] = not falhas
        d["Sinal"] = "COMPRAR" if not falhas else "não atende"
        d["Não atende em"] = ", ".join(falhas) if falhas else ""
        linhas.append(d)
    out = pd.DataFrame(linhas)
    return out.sort_values(["_n", "Score (0-100)"], ascending=[False, False]).reset_index(drop=True)


# ------------------------------------------------------------------- excel
def titulo(ws, texto, ncols, sub):
    ws["A1"] = texto
    ws["A1"].font = F_TIT
    ws.row_dimensions[1].height = 28
    for c in range(1, ncols + 1):
        ws.cell(1, c).fill = AZUL
    ws["A2"] = sub
    ws["A2"].font = F_NOTA
    ws.sheet_view.showGridLines = False


def escrever(ws, r0, df, fmts=None, larguras=None, cor_col=None):
    fmts = fmts or {}
    for j, c in enumerate(df.columns, 1):
        cel = ws.cell(r0, j, c)
        cel.font, cel.fill = F_CAB, AZUL
        cel.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        col = get_column_letter(j)
        larg = (larguras or {}).get(c, max(12, min(len(str(c)) + 3, 20)))
        if ws.column_dimensions[col].width < larg:
            ws.column_dimensions[col].width = larg
    ws.row_dimensions[r0].height = 30
    for i, (_, lin) in enumerate(df.iterrows(), 1):
        for j, c in enumerate(df.columns, 1):
            v = lin[c]
            if isinstance(v, float) and math.isnan(v):
                v = None
            if hasattr(v, "item"):
                v = v.item()
            cel = ws.cell(r0 + i, j, v)
            cel.font, cel.border = F_N, BORDA
            if c in fmts:
                cel.number_format = fmts[c]
            if isinstance(v, str) and len(v) > 38:
                cel.alignment = Alignment(wrap_text=True, vertical="top")
            if cor_col and c == cor_col and v in COR_SINAL:
                cel.fill, cel.font = COR_SINAL[v], F_B
    return r0 + len(df) + 2


def par_linha(ws, r, rotulo, valor, fmt=None, nota=None):
    ws.cell(r, 1, rotulo).font = F_N
    c = ws.cell(r, 2, valor)
    c.font = F_B
    if fmt:
        c.number_format = fmt
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
        return alt


# --------------------------------------------------------------- dashboard
NOMES = {
    "ITUB4": "Itaú Unibanco", "ITUB3": "Itaú Unibanco", "PETR4": "Petrobras", "PETR3": "Petrobras", "VALE3": "Vale",
    "BBAS3": "Banco do Brasil", "BBDC4": "Bradesco", "BBDC3": "Bradesco", "ABEV3": "Ambev", "WEGE3": "WEG",
    "B3SA3": "B3", "SANB11": "Santander Brasil", "BPAC11": "BTG Pactual", "ITSA4": "Itaúsa", "SUZB3": "Suzano",
    "RENT3": "Localiza", "LREN3": "Lojas Renner", "RADL3": "RaiaDrogasil", "EQTL3": "Equatorial", "CPLE6": "Copel",
    "TAEE11": "Taesa", "ELET3": "Eletrobras", "ELET6": "Eletrobras", "EGIE3": "Engie Brasil", "CMIG4": "Cemig",
    "SBSP3": "Sabesp", "VIVT3": "Vivo", "TIMS3": "TIM", "RAIL3": "Rumo", "PRIO3": "PRIO", "VBBR3": "Vibra Energia",
    "UGPA3": "Ultrapar", "CSAN3": "Cosan", "GGBR4": "Gerdau", "CSNA3": "CSN", "JBSS3": "JBS", "BRFS3": "BRF",
    "MGLU3": "Magazine Luiza", "HAPV3": "Hapvida", "RDOR3": "Rede D'Or", "KLBN11": "Klabin", "TOTS3": "Totvs",
    "CPFE3": "CPFL Energia", "CYRE3": "Cyrela", "MRVE3": "MRV", "DIRR3": "Direcional", "TEND3": "Tenda",
    "PLPL3": "Plano&Plano", "JHSF3": "JHSF", "PINE4": "Pine", "CVCB3": "CVC", "BBSE3": "BB Seguridade",
    "CXSE3": "Caixa Seguridade", "MXRF11": "Maxi Renda", "KNUQ11": "Kinea Unique HY CDI", "KNRI11": "Kinea Renda Imobiliária",
    "HGLG11": "CSHG Logística", "XPLG11": "XP Log", "XPML11": "XP Malls", "VISC11": "Vinci Shopping Centers",
    "HGRU11": "CSHG Renda Urbana", "BTLG11": "BTG Logística", "KNCR11": "Kinea Rendimentos", "KNIP11": "Kinea Índices de Preços",
    "CPTS11": "Capitânia Securities", "MCCI11": "Mauá Capital Recebíveis", "KNSC11": "Kinea Securities",
    "RBRR11": "RBR Rendimento High Grade", "RECR11": "REC Recebíveis", "VGHF11": "Valora Hedge Fund",
    "HGCR11": "CSHG Recebíveis", "IRDM11": "Iridium Recebíveis", "TRXF11": "TRX Real Estate", "BCFF11": "BTG Fundo de Fundos",
}
COR = {"rf": "0B5D57", "rf2": "2BA396", "rf3": "8FD3C9", "ac": "3B5BDB", "ac2": "93A5F0", "fi": "C96A12", "fi2": "EBB682"}


def _rem_partes(rem):
    """'120% do CDI' -> ('CDI', '120%'); 'IPCA + 6,5%' -> ('IPCA', '+ 6,5% a.a.')."""
    rem = str(rem or "")
    m = re.match(r"([\d.,]+)% do CDI", rem)
    if m:
        v = float(m.group(1).replace(".", "").replace(",", "."))
        if v <= 3:
            v *= 100
        return "CDI", f"{v:g}%".replace(".", ",")
    m = re.match(r"(IPCA|Selic) \+ ([\d.,]+)%", rem)
    if m:
        return m.group(1).upper() if m.group(1) == "IPCA" else "Selic", f"+ {m.group(2)}% a.a."
    m = re.match(r"([\d.,]+)% a\.a\.", rem)
    if m:
        return "Pré-fixado", f"{m.group(1)}% a.a."
    return rem or "n/d", ""


def _serie_rf(proj_row, hoje):
    """Valores em 1 mes e em 12, 24, ..., 120 meses. Se o robo ainda nao gerou 36+ meses, extrapola com o ritmo do 2o ano."""
    def g(c):
        v = proj_row.get(c) if proj_row is not None else None
        return float(v) if ok(v) else None
    um = g("1 mês")
    anos = [g(f"{12 * a} meses") for a in range(1, 11)]
    if anos[0] is None:
        return um, [hoje] * 10
    for i in range(1, 10):
        if anos[i] is None:
            if i == 1 or anos[i - 1] is None:
                anos[i] = anos[i - 1]
            else:
                anos[i] = anos[i - 1] * (anos[1] / anos[0] if anos[1] and anos[0] else 1)
    return um, anos


def construir_dashboard(wb, ctx):
    from openpyxl.chart import BarChart, DoughnutChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import DataPoint
    from openpyxl.worksheet.datavalidation import DataValidation as DV

    def rotulos(**kw):
        d_ = DataLabelList()
        d_.showVal = d_.showPercent = d_.showCatName = d_.showSerName = d_.showLegendKey = d_.showLeaderLines = False
        for k_, v_ in kw.items():
            setattr(d_, k_, v_)
        return d_

    def limpar(gr):
        """Titulo e legenda fora da area do grafico (senao o Excel sobrepoe o texto ao desenho)."""
        if gr.title is not None:
            gr.title.overlay = False
        if gr.legend is not None:
            gr.legend.overlay = False
        return gr

    p, par, pm, dash_in = ctx["p"], ctx["par"], ctx["pm"], ctx["dash_in"]
    perfil, h_aloc = ctx["perfil"], ctx["h_aloc"]
    ws = wb.create_sheet("Dashboard", 0)
    dd = wb.create_sheet("Dados Dashboard")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "0B5D57"
    for c in range(1, 13):
        ws.column_dimensions[get_column_letter(c)].width = 30 if c == 1 else 15.5
    titulo(ws, "DASHBOARD DA CARTEIRA", 12, f"Atualizado em {AGORA:%d/%m/%Y às %H:%M} (o robô regrava a cada 15 minutos). "
           f"Células amarelas são seletores: escolha na lista.")
    F_KPI = Font(name=ARIAL, size=18, bold=True, color="14202B")
    F_LAB = Font(name=ARIAL, size=9, color="556270")
    F_H2 = Font(name=ARIAL, size=13, bold=True, color="1F3864")
    CARD = PatternFill("solid", fgColor="F4F6F8")
    SEL_FONT = Font(name=ARIAL, size=10, bold=True, color="0B5D57")
    lista_h = '"' + ",".join(HORIZ_ROTULOS) + '"'

    def seletor(cel, valor, lista, msg="Escolha na lista"):
        c = ws[cel]
        c.value, c.font, c.fill = valor, SEL_FONT, ENTRADA
        c.alignment = Alignment(horizontal="center", vertical="center")
        dv = DV(type="list", formula1=lista, allow_blank=False, showErrorMessage=True, errorTitle="Valor inválido", error=msg)
        ws.add_data_validation(dv)
        dv.add(cel)

    # ------------------------------------------------ perfil
    ws["A4"] = "Perfil de risco"
    ws["A4"].font = F_B
    ws["A4"].alignment = Alignment(vertical="center")
    seletor(DASH_PERFIL, perfil, '"Conservador,Moderado,Arrojado,Personalizado"')
    ws["D4"] = ("Escolha aqui ou na aba Parametros: as duas são a mesma escolha. Os alvos de alocação mudam na hora, no Excel. "
                "Se digitar por cima na Parametros, o robô alinha o Dashboard na próxima atualização.")
    ws["D4"].font = F_NOTA
    ws.merge_cells("D4:L4")
    ws["D4"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[4].height = 30
    ws.merge_cells("A2:L2")

    # ------------------------------------------------ dados base
    rf = p["rf"] if not p["rf"].empty else pd.DataFrame()
    rv = p["rv"] if not p["rv"].empty else pd.DataFrame()
    proj = p["proj"]
    hoje_rf = [float(x) for x in rf["Valor líquido estimado (R$)"]] if len(rf) else []
    total_rf = sum(hoje_rf)
    linhas_rf = []
    for i, (_, l) in enumerate(rf.iterrows()):
        pr = None
        if not proj.empty:
            m_ = proj[proj["Investimento"] == l["Investimento"]]
            if len(m_):
                pr = m_.iloc[0]
        um, anos = _serie_rf(pr, hoje_rf[i])
        indexador, taxa = _rem_partes(l.get("Remuneração"))
        _v = l.get("Vencimento")
        _dias = l.get("Dias p/ vencer")
        linhas_rf.append({"nome": l["Investimento"], "tipo": l.get("Tipo") or "", "idx": indexador, "taxa": taxa,
                          "hoje": hoje_rf[i], "um": um if um is not None else hoje_rf[i], "anos": anos,
                          "inst": instituicao_de(l["Investimento"], l.get("Tipo")), "fgc": cobre_fgc(l["Investimento"], l.get("Tipo")),
                          "venc": _v.strftime("%d/%m/%Y") if hasattr(_v, "strftime") else (str(_v) if _v else ""),
                          "dias": int(_dias) if ok(_dias) else None})
    # linhas de renda variavel
    def bloco_rv(classe):
        out = []
        for _, l in (rv[rv["Classe"] == classe].iterrows() if len(rv) else []):
            qtd = float(l["Qtd"]) if ok(l.get("Qtd")) else 0.0
            pmed = l.get("Preço médio")
            valor = float(l["Valor atual (R$)"]) if ok(l.get("Valor atual (R$)")) else 0.0
            dy = float(l["DY 12m"]) if ok(l.get("DY 12m")) else 0.0
            out.append({"nome": NOMES.get(l["Ticker"], l["Ticker"]), "tk": l["Ticker"], "qtd": qtd,
                        "cot": float(l["Cotação"]) if ok(l.get("Cotação")) else None,
                        "pm": float(pmed) if ok(pmed) else None,
                        "inv": qtd * float(pmed) if ok(pmed) else None, "val": valor, "anual": dy * valor})
        return sorted(out, key=lambda x: -x["val"])
    acoes, fiis = bloco_rv("Ações"), bloco_rv("FIIs")
    # Preço de compra: o preço pago na compra (aba Compras). Com várias compras do mesmo ativo, média ponderada dos preços pagos.
    cp0 = p["compras"]
    for x in acoes + fiis:
        x["pc"], x["ncompras"] = x["pm"], 0
        if not cp0.empty:
            m_ = cp0[(cp0["Ticker"] == x["tk"]) & (cp0["Tipo"] == "Compra")]
            if len(m_):
                x["ncompras"] = len(m_)
                x["pc"] = float((m_["Quantidade"] * m_["Preço unitário (R$)"]).sum() / m_["Quantidade"].sum())
    total_ac, total_fi = sum(x["val"] for x in acoes), sum(x["val"] for x in fiis)
    total_rv = total_ac + total_fi

    # ------------------------------------------------ aba de dados
    titulo(dd, "DADOS DO DASHBOARD (gerados pelo robô; não edite)", 15,
           "Alimenta os seletores e os gráficos da aba Dashboard.")
    dd["A3"] = "Horizonte"
    dd["B3"] = "Anos"
    for c in ("A3", "B3"):
        dd[c].font = F_B
    for i, rot in enumerate(HORIZ_ROTULOS):
        dd.cell(4 + i, 1, rot)
        dd.cell(4 + i, 2, 1 if i == 0 else i + 1)
    H0, H1 = 4, 13      # linhas dos rotulos de horizonte
    dd["D3"], dd["E3"] = "Crescimento anual dos dividendos", f"=N(Parametros!$B${pm['div_cresc']})/100"
    dd["E3"].number_format = PC
    dd.column_dimensions["A"].width = 34
    # projecao da renda fixa por horizonte
    R0 = 17
    cab = ["Investimento", "Hoje", "1 mês"] + HORIZ_ROTULOS
    for j, c in enumerate(cab, 1):
        cel = dd.cell(R0, j, c)
        cel.font, cel.fill = F_CAB, AZUL
    for i, x in enumerate(linhas_rf):
        r = R0 + 1 + i
        dd.cell(r, 1, x["nome"])
        dd.cell(r, 2, x["hoje"])
        dd.cell(r, 3, x["um"])
        for k, v in enumerate(x["anos"]):
            dd.cell(r, 4 + k, v)
    rt = R0 + 1 + len(linhas_rf)
    dd.cell(rt, 1, "TOTAL").font = F_B
    for j in range(2, 14):
        col = get_column_letter(j)
        dd.cell(rt, j, f"=SUM({col}{R0 + 1}:{col}{rt - 1})" if linhas_rf else 0)
    for r in range(R0 + 1, rt + 1):
        for j in range(2, 14):
            dd.cell(r, j).number_format = RS
    # dividendos
    D0 = rt + 4
    for j, c in enumerate(["Ticker", "Classe", "Valor atual (R$)", "Dividendos 12 meses (R$)", "Média mensal (R$)"], 1):
        cel = dd.cell(D0, j, c)
        cel.font, cel.fill = F_CAB, AZUL
    drows = {}
    r = D0 + 1
    for cl, lst in (("Ações", acoes), ("FIIs", fiis)):
        for x in lst:
            dd.cell(r, 1, x["tk"])
            dd.cell(r, 2, cl)
            dd.cell(r, 3, x["val"]).number_format = RS
            dd.cell(r, 4, x["anual"]).number_format = RS
            dd.cell(r, 5, f"=D{r}/12").number_format = RS
            drows[x["tk"]] = r
            r += 1
    d_ini, d_fim = D0 + 1, max(r - 1, D0 + 1)
    # melhores
    M0 = r + 3

    def fator(sel_cell):
        n = f"INDEX('Dados Dashboard'!$B${H0}:$B${H1},MATCH({sel_cell},'Dados Dashboard'!$A${H0}:$A${H1},0))"
        g = "'Dados Dashboard'!$E$3"
        return f"IF({g}=0,{n},((1+{g})^{n}-1)/{g})"

    def idx_h(sel_cell):
        return f"MATCH({sel_cell},'Dados Dashboard'!$A${H0}:$A${H1},0)"

    # ------------------------------------------------ cartoes
    def cartao(c0, c1, r0, rotulo, valor, sub, fmt=RS, cor="14202B"):
        for rr in range(r0, r0 + 3):
            for cc in range(c0, c1 + 1):
                ws.cell(rr, cc).fill = CARD
        ws.merge_cells(start_row=r0, start_column=c0, end_row=r0, end_column=c1)
        ws.merge_cells(start_row=r0 + 1, start_column=c0, end_row=r0 + 1, end_column=c1)
        ws.merge_cells(start_row=r0 + 2, start_column=c0, end_row=r0 + 2, end_column=c1)
        a = ws.cell(r0, c0, rotulo)
        a.font, a.alignment = F_LAB, Alignment(vertical="center", indent=1)
        b = ws.cell(r0 + 1, c0, valor)
        b.font, b.number_format = Font(name=ARIAL, size=18, bold=True, color=cor), fmt
        b.alignment = Alignment(vertical="center", indent=1, horizontal="left")
        s_ = ws.cell(r0 + 2, c0, sub)
        s_.font, s_.alignment = F_LAB, Alignment(vertical="center", indent=1)

    for rr in (7, 11, 15):
        ws.row_dimensions[rr].height = 18
        ws.row_dimensions[rr + 1].height = 30
        ws.row_dimensions[rr + 2].height = 18
    ws["A6"] = "VISÃO GERAL"
    ws["A6"].font = F_H2
    cartao(1, 3, 7, "Patrimônio total", "=A12+A16", '=ROUND(A12/MAX(A8,0.01)*100,1)&"% renda fixa · "&ROUND(A16/MAX(A8,0.01)*100,1)&"% renda variável"')
    # renda fixa
    t_rf = f"'Dados Dashboard'!"
    cartao(1, 3, 11, "Renda fixa", total_rf, f"{len(linhas_rf)} aplicação(ões) · valor líquido estimado", cor="0B5D57")
    cartao(4, 6, 11, "Renda fixa: projeção para o próximo mês", f"={t_rf}C{rt}",
           f'="+R$ "&ROUND(D12-A12,2)&" líquidos de IR (estimativa)"')
    cartao(7, 9, 11, "Renda fixa: projeção para daqui a 12 meses", f"={t_rf}D{rt}",
           f'="+"&ROUND((G12/MAX(A12,0.01)-1)*100,1)&"% líquido · cenário Focus"')
    cartao(10, 12, 11, "Renda fixa: projeção para daqui a", f"=INDEX({t_rf}D{rt}:M{rt},{idx_h(DASH_H_RF)})",
           "líquido de IR · cenário Focus, sem novos aportes")
    ws.unmerge_cells("J11:L11")
    ws.merge_cells("J11:K11")
    ws["J11"].alignment = Alignment(vertical="center", indent=1, horizontal="right")
    seletor(DASH_H_RF, dash_in.get("rf_card", "2 anos"), lista_h)
    # renda variavel
    tem_rv = bool(acoes or fiis)
    cartao(1, 3, 15, "Renda variável", total_rv, f"{len(acoes)} ação(ões) · {len(fiis)} FII(s)" if tem_rv else
           "nenhum ativo cadastrado ainda", cor="3B5BDB")
    soma_mes = f"SUM({t_rf}E{d_ini}:E{d_fim})" if tem_rv else "0"
    soma_ano = f"SUM({t_rf}D{d_ini}:D{d_fim})" if tem_rv else "0"
    cartao(4, 6, 15, "Dividendos estimados: próximo mês", f"={soma_mes}",
           "média mensal (ações pagam em parcelas irregulares)")
    cartao(7, 9, 15, "Dividendos estimados: próximos 12 meses", f"={soma_ano}",
           '=ROUND(G16/MAX(A16,0.01)*100,1)&"% sobre o valor atual"')
    cartao(10, 12, 15, "Dividendos estimados: próximos", f"={soma_ano}*{fator(DASH_H_DV)}" if tem_rv else 0,
           "soma acumulada; crescimento anual na aba Parametros")
    ws.unmerge_cells("J15:L15")
    ws.merge_cells("J15:K15")
    ws["J15"].alignment = Alignment(vertical="center", indent=1, horizontal="right")
    seletor(DASH_H_DV, dash_in.get("dv_card", "2 anos"), lista_h)

    # ------------------------------------------------ alocacao
    ws["A19"] = "ALOCAÇÃO: HOJE × ALVO DO PERFIL"
    ws["A19"].font = F_H2
    for j, c in enumerate(["Classe", "Hoje", "Alvo do perfil", "Diferença (R$)", "Destino do aporte (R$)"], 1):
        cel = ws.cell(20, j, c)
        cel.font, cel.fill = F_CAB, AZUL
        cel.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[20].height = 28
    for i, nome in enumerate(["Renda fixa", "Ações", "FIIs"]):
        rr, src = 21 + i, h_aloc + 1 + i
        ws.cell(rr, 1, nome).font = F_B
        for j, col, fmt in ((2, "C", PC), (3, "D", PC), (4, "E", RS), (5, "F", RS)):
            cel = ws.cell(rr, j, f"='Plano de Acao'!{col}{src}")
            cel.number_format, cel.font, cel.border = fmt, F_N, BORDA
    ws["A24"] = "Aporte disponível agora (R$)"
    ws["A24"].font = F_B
    ws["A24"].alignment = Alignment(vertical="center")
    ws[DASH_APORTE] = float(par.get("aporte") or 0)
    ws[DASH_APORTE].font, ws[DASH_APORTE].fill, ws[DASH_APORTE].number_format = SEL_FONT, ENTRADA, RS
    ws[DASH_APORTE].alignment = Alignment(horizontal="center", vertical="center")
    dv_ap = DV(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=False, showErrorMessage=True,
               errorTitle="Valor inválido", error="Digite um valor em reais, zero ou maior.")
    ws.add_data_validation(dv_ap)
    dv_ap.add(DASH_APORTE)
    ws["C24"] = "Digite aqui ou na aba Parametros: é a mesma célula. A tabela e o Plano de Acao se recalculam na hora."
    ws["C24"].font = F_NOTA
    ws["A25"] = "Diferença positiva = falta para o alvo. O aporte é dividido entre as classes que faltam."
    ws["A25"].font = F_NOTA
    ws.row_dimensions[24].height = 22
    ch = BarChart()
    ch.type = "col"
    ch.title = "Alocação: hoje × alvo"
    ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=20, max_row=23), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=21, max_row=23))
    ch.y_axis.number_format = "0%"
    ch.y_axis.majorGridlines = None
    ch.x_axis.delete = ch.y_axis.delete = False
    ch.series[0].graphicalProperties.solidFill = "0B5D57"
    ch.series[1].graphicalProperties.solidFill = "14202B"
    ch.legend.position = "b"
    ch.y_axis.scaling.min = 0
    ch.y_axis.scaling.max = 1
    ch.y_axis.majorUnit = 0.2
    ch.gapWidth = 80
    ch.width, ch.height = 15, 7
    ws.add_chart(limpar(ch), "G19")

    # ------------------------------------------------ carteira
    r = 34
    ws.cell(r, 1, "1 · MINHA CARTEIRA ATUAL").font = F_H2
    r += 2

    def cab_tabela(r, cols):
        for j, c in enumerate(cols, 1):
            cel = ws.cell(r, j, c)
            cel.font, cel.fill = F_CAB, AZUL
            cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 42

    def donut(titulo_, r_cab, r_ini, r_fim, col_nome, col_val, ancora, cores):
        d = DoughnutChart()
        d.title = titulo_
        d.holeSize = 55
        d.add_data(Reference(ws, min_col=col_val, min_row=r_ini, max_row=r_fim), titles_from_data=False)
        d.set_categories(Reference(ws, min_col=col_nome, min_row=r_ini, max_row=r_fim))
        d.dataLabels = rotulos(showPercent=True)
        ser = d.series[0]
        for k in range(r_fim - r_ini + 1):
            pt = DataPoint(idx=k)
            pt.graphicalProperties.solidFill = cores[k % len(cores)]
            ser.dPt.append(pt)
        d.legend.position = "b"
        d.width, d.height = 12.5, 8.5
        ws.add_chart(limpar(d), ancora)

    # ---- renda fixa
    ws.cell(r, 1, "Renda fixa").font = F_B
    ws.cell(r, 3, "dados reais · valores líquidos de IR, estimados").font = F_NOTA
    r += 1
    ws.cell(r, 1, "Horizonte (renda fixa)").font = F_N
    sel_rf = f"$B${r}"
    seletor(f"B{r}", dash_in.get("rf_tab", "12 meses"), lista_h)
    r += 1
    cab_tabela(r, ["Investimento", "Tipo", "Indexador", "Taxa atual", "Valor hoje (R$)", "Projeção próx. mês (R$)", "Projeção no horizonte (R$)"])
    r_cab = r
    ini = r + 1
    for i, x in enumerate(linhas_rf):
        rr = ini + i
        dr = R0 + 1 + i
        vals = [x["nome"], x["tipo"], x["idx"], x["taxa"], x["hoje"], x["um"],
                f"=INDEX('Dados Dashboard'!D{dr}:M{dr},{idx_h(sel_rf)})"]
        for j, v in enumerate(vals, 1):
            cel = ws.cell(rr, j, v)
            cel.font, cel.border = F_N, BORDA
            if j >= 5:
                cel.number_format = RS
            if j in (2, 3, 4):
                cel.alignment = Alignment(horizontal="center")
    fim = ini + len(linhas_rf) - 1
    rt_ = fim + 1
    ws.cell(rt_, 1, "Total")
    for j in (5, 6, 7):
        col = get_column_letter(j)
        ws.cell(rt_, j, f"=SUM({col}{ini}:{col}{fim})" if linhas_rf else 0).number_format = RS
    for j in range(1, 8):
        ws.cell(rt_, j).font = F_B
        ws.cell(rt_, j).border = Border(top=Side(style="thin", color="7F7F7F"))
    if linhas_rf:
        donut("Renda fixa por investimento", r_cab, ini, fim, 1, 5, f"I{r_cab - 2}", [COR["rf"], COR["rf2"], COR["rf3"], "5FB9AE", "BFE6E0"])
    else:
        ws.cell(ini, 1, "Nenhum investimento em Minha_Carteira.xlsx (aba Renda Fixa).").font = F_N
    r = max(rt_ + 3, r_cab - 2 + 17)

    # ---- renda variavel
    def bloco(r, titulo_, lst, chave, cores, cor_chave):
        ws.cell(r, 1, titulo_).font = F_B
        ws.cell(r, 3, "dividendos estimados pelo histórico de 12 meses do ativo" if lst else "").font = F_NOTA
        r += 1
        if not lst:
            ws.cell(r, 1, f"Você ainda não tem {'ações' if chave == 'ac' else 'FIIs'} cadastrados na aba Renda Variavel de Minha_Carteira.xlsx. "
                          "Quando cadastrar, esta área mostra o gráfico e a tabela.").font = F_NOTA
            return r + 3
        ws.cell(r, 1, f"Horizonte ({'ações' if chave == 'ac' else 'FIIs'})").font = F_N
        sel = f"$B${r}"
        seletor(f"B{r}", dash_in.get(f"{chave}_tab", "12 meses"), lista_h)
        r += 1
        cab_tabela(r, ["Empresa" if chave == "ac" else "Fundo", "Ticker", "Quantidade", "Preço de compra (R$)", "Cotação atual (R$)", "Investido (R$)", "Valor atual (R$)",
                       "Dividendo próx. mês (R$)", "Dividendo acumulado no horizonte (R$)"])
        r_cab = r
        ini = r + 1
        for i, x in enumerate(lst):
            rr = ini + i
            dr = drows[x["tk"]]
            vals = [x["nome"], x["tk"], x["qtd"], x["pc"], x["cot"], x["inv"], x["val"], f"='Dados Dashboard'!E{dr}",
                    f"='Dados Dashboard'!D{dr}*{fator(sel)}"]
            for j, v in enumerate(vals, 1):
                cel = ws.cell(rr, j, v)
                cel.font, cel.border = F_N, BORDA
                if j == 3:
                    cel.number_format = N0
                if j >= 4:
                    cel.number_format = RS
                if j == 2:
                    cel.alignment = Alignment(horizontal="center")
        fim = ini + len(lst) - 1
        ws.cell(fim + 1, 1, "Total")
        for j in (3, 6, 7, 8, 9):
            col = get_column_letter(j)
            ws.cell(fim + 1, j, f"=SUM({col}{ini}:{col}{fim})").number_format = N0 if j == 3 else RS
        for j in range(1, 10):
            ws.cell(fim + 1, j).font = F_B
            ws.cell(fim + 1, j).border = Border(top=Side(style="thin", color="7F7F7F"))
        ws.cell(fim + 2, 1, "Preço de compra = preço pago, buscado em Minhas compras e vendas. Com mais de uma compra do mesmo ativo, "
                            "mostra a média dos preços pagos. Investido = quantidade × preço médio com taxas.").font = F_NOTA
        donut("Ações por valor atual" if chave == "ac" else "FIIs por valor atual", r_cab, ini, fim, 1, 7, f"K{r_cab - 2}", cores)
        return max(fim + 4, r_cab - 2 + 17)

    r = bloco(r, "Renda variável: ações", acoes, "ac", [COR["ac"], COR["ac2"], "5C7CE6", "BAC6F5", "2F46A8"], "ac")
    r = bloco(r, "Renda variável: fundos imobiliários", fiis, "fi", [COR["fi"], COR["fi2"], "DE8A3A", "F3D2B0", "9A4E0C"], "fi")

    # ---- historico de compras
    lotes_html = []
    cp = p["compras"]
    ws.cell(r, 1, "Minhas compras e vendas").font = F_B
    ws.cell(r, 3, "uma linha por operação, da aba Compras de Minha_Carteira.xlsx").font = F_NOTA
    r += 1
    if cp.empty:
        ws.cell(r, 1, "Nenhuma operação registrada ainda. Lance cada compra na aba Compras de Minha_Carteira.xlsx: "
                      "o robô calcula o preço médio e mostra o histórico aqui.").font = F_NOTA
        r += 3
    else:
        classe_tk = {}
        for df_, nm_ in ((ctx["ca"], "Ação"), (ctx["cf"], "Fundo")):
            if df_ is not None and not df_.empty:
                for t_ in df_["Ticker"]:
                    classe_tk[str(t_)] = nm_
        if len(rv):
            for t_, c_ in zip(rv["Ticker"], rv["Classe"]):
                if c_ in ("Ações", "FIIs"):
                    classe_tk[str(t_)] = "Fundo" if c_ == "FIIs" else "Ação"
        cab_tabela(r, ["Empresa / fundo", "Ticker", "Carteira", "Data", "Tipo", "Quantidade", "Preço unitário (R$)", "Taxas (R$)",
                       "Valor da operação (R$)", "Cotação atual (R$)", "Resultado (R$)", "Resultado (%)"])
        RSV = 'R$ #,##0.00;[Red]-R$ #,##0.00'
        ini = r + 1
        for i, (_, l) in enumerate(cp.iterrows()):
            rr = ini + i
            tk = str(l["Ticker"])
            d_ = l["Data"]
            cart_ = classe_tk.get(tk) or ("Fundo" if tk.endswith("11") and tk[:4] not in ("TAEE", "KLBN", "SANB", "ALUP", "BPAC") else "Ação")
            vals = [NOMES.get(tk, tk), tk, cart_, d_, l["Tipo"], l["Quantidade"], l["Preço unitário (R$)"], l["Taxas (R$)"],
                    l["Valor da operação (R$)"], l["Cotação atual (R$)"], l["Resultado (R$)"], l["Resultado (%)"]]
            vals = [None if (isinstance(v, float) and v != v) else v for v in vals]
            for j, v in enumerate(vals, 1):
                cel = ws.cell(rr, j, v)
                cel.font, cel.border = F_N, BORDA
                cel.number_format = {4: "DD/MM/YYYY", 6: N0, 7: RS, 8: RS, 9: RS, 10: RS, 11: RSV, 12: "0.0%;[Red]-0.0%"}.get(j, "General")
                if j in (2, 3, 4, 5):
                    cel.alignment = Alignment(horizontal="center")
            lotes_html.append({"data": d_.strftime("%d/%m/%Y") if hasattr(d_, "strftime") else str(d_), "tk": tk,
                               "nome": NOMES.get(tk, tk), "cart": cart_, "tipo": l["Tipo"], "qtd": vals[5], "preco": vals[6],
                               "taxas": vals[7], "valor": vals[8], "cot": vals[9], "res": vals[10], "pct": vals[11]})
        r = ini + len(cp) + 2

    # ------------------------------------------------ melhores
    ws.cell(r, 1, "2 · MELHORES OPÇÕES AGORA, PELOS SEUS CRITÉRIOS").font = F_H2
    ws.cell(r + 1, 1, "Nota 0 a 100 do ranking do robô, entre os ativos que mais atendem à aba Parametros. Não é recomendação de compra.").font = F_NOTA
    r += 3
    rm = M0
    dd.cell(rm - 1, 1, "MELHORES (gráficos do Dashboard)").font = F_B

    def lista_melhores(df, nome_cols):
        out = []
        if df is None or df.empty:
            return out
        for _, l in df.head(5).iterrows():
            dy = l.get("DY 12m")
            rot = f"{l['Ticker']} · DY {fbr(dy * 100, 1)}%" if ok(dy) else str(l["Ticker"])
            out.append((rot, float(l["Score (0-100)"]) / 100 if ok(l.get("Score (0-100)")) else 0))
        return out
    mac, mfi = lista_melhores(ctx["ca"], None), lista_melhores(ctx["cf"], None)
    tk_ac = [str(t) for t in ctx["ca"]["Ticker"].head(5)] if not ctx["ca"].empty else []
    tk_fi = [str(t) for t in ctx["cf"]["Ticker"].head(5)] if not ctx["cf"].empty else []
    fichas = {}
    for df_, cls_, cols_ in ((ctx["ca"], "ac", COLS_AC), (ctx["cf"], "fi", cols_fi(ctx["cf"]))):
        if df_ is None or df_.empty:
            continue
        for pos_, (_, l_) in enumerate(df_.head(40).iterrows(), 1):
            fichas[f"{cls_}:{l_['Ticker']}"] = {"nome": NOMES.get(l_["Ticker"], ""), "pos": pos_,
                                                "campos": ficha_linha(l_, cols_)}
    # renda fixa
    mrf = []
    cdi = ctx.get("cdi")
    ref = p["rf_ref"]
    nomes_tem = {}
    for x in linhas_rf:
        if x["idx"] == "CDI" and cdi:
            pct = float(x["taxa"].replace("%", "").replace(",", "."))
            nomes_tem[pct] = x["tipo"] or "RDB"
    if not ref.empty:
        col17 = next((c for c in ref.columns if "1 a 2 anos" in c), None)
        for _, l in ref.iterrows():
            if col17 and ok(l[col17]):
                nome = str(l["Aplicação"]).split(" (aprox")[0]
                m_ = re.search(r"CDB ([\d.,]+)% do CDI", nome)
                if m_ and float(m_.group(1).replace(",", ".")) in nomes_tem:
                    nome += " (você tem)"
                    nomes_tem.pop(float(m_.group(1).replace(",", ".")))
                mrf.append((nome, float(l[col17])))
        for pct, tp in nomes_tem.items():
            mrf.append((f"{tp} {pct:g}% do CDI (você tem)".replace(".", ","), cdi / 100 * pct / 100 * (1 - 0.175)))
    ip, ipc = ctx.get("ipca_mais"), ctx.get("ipca_fut")
    if ip is not None and ipc is not None:
        mrf.append(("Tesouro IPCA+ (estimativa)", ((1 + ip) * (1 + ipc / 100) - 1) * (1 - 0.15)))
    mrf = sorted(mrf, key=lambda t: -t[1])[:8]

    def tabela_dados(r_, titulo_, lst, col):
        dd.cell(r_, col, titulo_).font = F_B
        for i, (a, b) in enumerate(lst):
            dd.cell(r_ + 1 + i, col, a)
            c_ = dd.cell(r_ + 1 + i, col + 1, b)
            c_.number_format = "0.0%"
        return r_ + 1, r_ + len(lst)
    a0, a1 = tabela_dados(rm, "Ações (nota)", mac, 1)
    f0, f1 = tabela_dados(rm, "FIIs (nota)", mfi, 4)
    g0, g1 = tabela_dados(rm, "Renda fixa (retorno líquido a.a., 1 a 2 anos)", mrf, 7)

    def barras(titulo_, d0, d1, col, ancora, cor, fmt, larg=13.2):
        if d1 < d0:
            return
        b = BarChart()
        b.type = "bar"
        b.title = titulo_
        b.add_data(Reference(dd, min_col=col + 1, min_row=d0, max_row=d1), titles_from_data=False)
        b.set_categories(Reference(dd, min_col=col, min_row=d0, max_row=d1))
        b.x_axis.scaling.orientation = "maxMin"
        b.x_axis.delete = b.y_axis.delete = False
        b.y_axis.number_format = "0%"
        b.y_axis.scaling.min = 0
        b.y_axis.majorGridlines = None
        b.legend = None
        b.dataLabels = rotulos(showVal=True, numFmt=fmt)
        b.dataLabels.position = "outEnd"
        b.y_axis.scaling.max = 1 if fmt == "0%" else None
        b.y_axis.delete = True
        b.series[0].graphicalProperties.solidFill = cor
        b.gapWidth = 60
        b.width, b.height = larg, 8.5
        ws.add_chart(limpar(b), ancora)
    barras("Ações: nota (0 a 100)", a0, a1, 1, f"A{r}", COR["ac"], "0%", 17)
    barras("FIIs: nota (0 a 100)", f0, f1, 4, f"F{r}", COR["fi"], "0%", 17)
    from openpyxl.worksheet.hyperlink import Hyperlink
    ws.cell(r + 17, 1, "Clique no ativo para ver todos os dados dele (aba Candidatos):").font = F_NOTA
    for nome_aba, tks_, col0 in (("Candidatos Acoes", tk_ac, 1), ("Candidatos FIIs", tk_fi, 6)):
        for i_, tk_ in enumerate(tks_):
            c_ = ws.cell(r + 18, col0 + i_, tk_)
            c_.hyperlink = Hyperlink(ref=c_.coordinate, location=f"'{nome_aba}'!B{5 + i_}", display=tk_)
            c_.font = Font(name="Arial", size=10, color="0563C1", underline="single", bold=True)
    r += 19
    barras("Renda fixa: retorno líquido ao ano (1 a 2 anos, após IR)", g0, g1, 7, f"A{r}", COR["rf"], "0.0%", 17)
    if not mac:
        ws.cell(r, 1, "Sem dados de ações no painel nesta atualização.").font = F_NOTA
    r += 19
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    ws.row_dimensions[r].height = 42
    ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(r, 1, ("Notas e retornos são cálculos mecânicos sobre dados públicos e os critérios editáveis da aba Parametros; "
                   "projeções partem do Boletim Focus e podem não se confirmar. " + AVISO)).font = F_NOTA
    dd.sheet_view.showGridLines = True
    try:
        extras = construir_extras(wb, ctx, linhas_rf, acoes, fiis, total_rf, total_ac, total_fi)
    except Exception as e:  # noqa
        print("[extras] falha:", repr(e)[:200], flush=True)
        extras = None
    ws.sheet_view.zoomScale = 90
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return {
        "gerado": f"{AGORA:%d/%m/%Y às %H:%M}", "perfil": perfil, "horizontes": HORIZ_ROTULOS,
        "alvos": {k: list(v) for k, v in par["tabela_alvos"].items()},
        "aporte": float(par.get("aporte") or 0), "div_cresc": float(par.get("div_cresc") or 0),
        "rf": linhas_rf, "ac": acoes, "fi": fiis, "compras": lotes_html,
        "melhores_ac": [list(t) + [tk] for t, tk in zip(mac, tk_ac)], "melhores_fi": [list(t) + [tk] for t, tk in zip(mfi, tk_fi)],
        "melhores_rf": [list(t) for t in mrf], "fichas": fichas,
        "x": extras, "aviso": AVISO}


# ============================================================ EXTRAS DO DASHBOARD
# metas, historico e benchmarks, "o que fazer hoje", proventos, mapa de risco
MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
HIST_COLS = ["Data", "Patrimônio (R$)", "Renda fixa (R$)", "Ações (R$)", "FIIs (R$)", "Aplicado total (R$)",
             "CDI (% a.a.)", "IPCA 12m (%)", "Ibovespa", "IFIX", "Dólar (R$)"]
_TIPOS_INST = r"(?:RDB|CDB|LCI|LCA|LC|LF|LIG|CRI|CRA|CDCA|Debênture|Debenture|Poupança|Tesouro(?: \w+)?)"


def instituicao_de(nome, tipo=""):
    """'RDB Nubank (Poupança)' -> 'Nubank'. Titulos do Tesouro -> 'Tesouro Nacional'."""
    if "tesouro" in str(nome).lower() or "tesouro" in str(tipo).lower():
        return "Tesouro Nacional"
    n = re.sub(r"\(.*?\)", "", str(nome)).strip()
    n = re.sub(rf"^(?:{_TIPOS_INST})\s+", "", n, flags=re.I).strip()
    return n or str(nome)


def cobre_fgc(nome, tipo):
    t = (str(tipo) + " " + str(nome)).upper()
    if any(x in t for x in ("TESOURO", "DEBÊNTURE", "DEBENTURE")) or re.search(r"\b(CRI|CRA)\b", t):
        return False
    return bool(re.search(r"\b(RDB|CDB|LCI|LCA|LC|LF|POUPAN[CÇ]A)\b", t))


def ler_historico(caminho):
    """Historico diario e fotos do dia anterior, guardados no proprio Plano_de_Acao.xlsx."""
    linhas, snaps = [], {}
    try:
        if Path(caminho).exists():
            wb = load_workbook(caminho, data_only=True)
            if "Historico" in wb.sheetnames:
                for lin in wb["Historico"].iter_rows(min_row=2, max_col=len(HIST_COLS), values_only=True):
                    if lin[0] and isinstance(lin[1], (int, float)):
                        linhas.append(list(lin))
            if "Snapshots" in wb.sheetnames:
                import json as _j
                for d, js in wb["Snapshots"].iter_rows(min_row=2, max_col=2, values_only=True):
                    try:
                        snaps[str(d)] = _j.loads(js)
                    except Exception:  # noqa
                        pass
            wb.close()
    except Exception:  # noqa
        pass
    return linhas, snaps


def _num_or_none(v):
    return float(v) if ok(v) else None


def serie_desempenho(linhas):
    """Indices base 100: carteira (retorno ponderado pelo tempo, sem contar aportes), CDI, IPCA, Ibovespa, IFIX."""
    if len(linhas) < 2:
        return None
    ds = [dt.date.fromisoformat(str(l[0])[:10]) for l in linhas]
    cart, cdi, ipca, ibov, ifix = [100.0], [100.0], [100.0], [100.0], [100.0]
    for i in range(1, len(linhas)):
        a, b = linhas[i - 1], linhas[i]
        v0, v1 = _num_or_none(a[1]), _num_or_none(b[1])
        ap0, ap1 = _num_or_none(a[5]), _num_or_none(b[5])
        fluxo = (ap1 - ap0) if ap0 is not None and ap1 is not None else 0.0
        r = ((v1 - fluxo) / v0 - 1) if v0 and v1 is not None else 0.0
        cart.append(cart[-1] * (1 + r))
        try:
            import numpy as _np
            du = int(_np.busday_count(ds[i - 1], ds[i]))
        except Exception:  # noqa
            du = max(0, (ds[i] - ds[i - 1]).days * 5 // 7)
        c0 = _num_or_none(a[6])
        cdi.append(cdi[-1] * ((1 + c0 / 100) ** (du / 252) if c0 else 1))
        i1 = _num_or_none(b[7])
        ipca.append(ipca[-1] * ((1 + i1 / 100) ** ((ds[i] - ds[i - 1]).days / 365) if i1 else 1))
        for lst, k in ((ibov, 8), (ifix, 9)):
            x0, x1 = _num_or_none(a[k]), _num_or_none(b[k])
            lst.append(lst[-1] * (x1 / x0 if x0 and x1 else 1))
    return {"datas": [d.strftime("%d/%m") for d in ds], "carteira": cart, "cdi": cdi, "ipca": ipca,
            "ibov": ibov, "ifix": ifix, "patrimonio": [_num_or_none(l[1]) for l in linhas],
            "aplicado": [_num_or_none(l[5]) for l in linhas], "inicio": ds[0].strftime("%d/%m/%Y")}


def calendario_proventos(p_prov, ativos, hoje):
    """Proximos 12 meses de proventos: o que cada ativo pagou nos ultimos 12 meses x quantidade atual.
    Cada item traz o dia estimado (data 'com' do ano passado, repetida neste ano)."""
    meses = []
    y, m = hoje.year, hoje.month
    for i in range(12):
        meses.append((y + (m - 1 + i) // 12, (m - 1 + i) % 12 + 1))
    tot = {k: {} for k in meses}          # mes -> ticker -> [valor, [dias]]
    usou_dy = []

    def soma(k, tk, v, dia):
        d = tot[k].setdefault(tk, [0.0, []])
        d[0] += v
        if dia and dia not in d[1]:
            d[1].append(dia)
    for a in ativos:
        tk, qtd = a["tk"], a["qtd"]
        ev = []
        if p_prov is not None and not p_prov.empty:
            sub = p_prov[p_prov["Ticker"] == tk]
            for d_, v_ in zip(sub["Data"], sub["Valor (R$/cota)"]):
                try:
                    d0 = dt.date.fromisoformat(str(d_)[:10])
                except ValueError:
                    continue
                if (hoje - d0).days <= 365 and d0 <= hoje:
                    ev.append((d0, float(v_)))
        if ev:
            mensal = len({(d.year, d.month) for d, _ in ev}) >= 9
            if mensal:
                med = sum(v for _, v in ev) / 12
                dias = sorted(d.day for d, _ in ev)
                dia = dias[len(dias) // 2]
                for k in meses:
                    soma(k, tk, med * qtd, dia)
            else:
                for d, v in ev:
                    k = (d.year + 1, d.month)
                    if k in tot:
                        soma(k, tk, v * qtd, d.day)
        elif a.get("anual"):
            for k in meses:
                soma(k, tk, a["anual"] / 12, None)
            usou_dy.append(tk)
    cal = []
    for k, v in tot.items():
        itens = [[t, x[0], "/".join(f"{d:02d}" for d in sorted(x[1])) + f"/{k[1]:02d}" if x[1] else ""]
                 for t, x in sorted(v.items(), key=lambda z: -z[1][0])]
        cal.append({"mes": f"{MESES[k[1] - 1]}/{str(k[0])[2:]}", "total": sum(x[0] for x in v.values()), "itens": itens})
    return cal, usou_dy


def construir_extras(wb, ctx, linhas_rf, acoes, fiis, total_rf, total_ac, total_fi):
    import json as _j
    p, par = ctx["p"], ctx["par"]
    m = p["macro"]
    hoje = AGORA.date()
    pat = total_rf + total_ac + total_fi
    sinais, destino = ctx.get("sinais", []), ctx["destino"]
    ca, cf = ctx["ca"], ctx["cf"]

    # ---------- aplicado total (para separar aporte de rentabilidade)
    aplicado = None
    try:
        rs = p.get("resumo")
        if rs is not None and not rs.empty:
            lt = rs[rs["Classe"].astype(str).str.upper() == "TOTAL"]
            if len(lt):
                aplicado = float(lt.iloc[0]["Investido (R$)"])
    except Exception:  # noqa
        pass
    if aplicado is None:
        aplicado = pat

    # ---------- historico
    hist, snaps = ler_historico(destino)
    iso = hoje.isoformat()
    nova = [iso, pat, total_rf, total_ac, total_fi, aplicado, m.get("CDI (% a.a.)"), m.get("IPCA 12 meses (%)"),
            m.get("Ibovespa (pontos)"), m.get("IFIX (pontos)"), m.get("Dólar comercial (R$)")]
    ant_hist = [l for l in hist if str(l[0])[:10] < iso]
    hist = [l for l in hist if str(l[0])[:10] != iso] + [nova]
    hist = sorted(hist, key=lambda l: str(l[0]))[-800:]
    desempenho = serie_desempenho(hist)

    # ---------- fotos (mudancas desde ontem)
    aprov_ac = [str(t) for t in ca[ca["_todos"]]["Ticker"]] if not ca.empty else []
    aprov_fi = [str(t) for t in cf[cf["_todos"]]["Ticker"]] if not cf.empty else []
    al_df = p.get("alertas")
    alertas_hoje = [str(a) for a in al_df["Alerta"]] if al_df is not None and not al_df.empty else []
    snap = {"aprov_ac": aprov_ac, "aprov_fi": aprov_fi, "sinais": {s["Ativo"]: s["Sinal"] for s in sinais},
            "alertas": alertas_hoje, "selic": m.get("Selic meta (% a.a.)"), "cdi": m.get("CDI (% a.a.)"),
            "dolar": m.get("Dólar comercial (R$)"), "ibov": m.get("Ibovespa (pontos)")}
    snaps[iso] = snap
    datas_ant = sorted(d for d in snaps if d < iso)
    prev = snaps.get(datas_ant[-1]) if datas_ant else None
    prev_data = datas_ant[-1] if datas_ant else None

    mud = []
    if ant_hist:
        a = ant_hist[-1]
        dif = pat - (a[1] or 0)
        fl = aplicado - (a[5] or 0)
        rent = dif - fl
        d_ant = dt.date.fromisoformat(str(a[0])[:10]).strftime("%d/%m")
        txt = f"Patrimônio {'subiu' if dif >= 0 else 'caiu'} R$ {fbr(abs(dif))} desde {d_ant}"
        if abs(fl) >= 0.01:
            txt += f" (aportes/resgates de R$ {fbr(fl)}; rentabilidade de R$ {fbr(rent)})"
        mud.append(txt + ".")
    if prev:
        for tk in aprov_ac + aprov_fi:
            if tk not in prev.get("aprov_ac", []) + prev.get("aprov_fi", []):
                mud.append(f"{tk} passou a atender a todos os critérios de compra.")
        for tk in prev.get("aprov_ac", []) + prev.get("aprov_fi", []):
            if tk not in aprov_ac + aprov_fi:
                mud.append(f"{tk} deixou de atender a todos os critérios.")
        for tk, s in snap["sinais"].items():
            s0 = prev.get("sinais", {}).get(tk)
            if s0 and s0 != s:
                mud.append(f"Sinal de {tk} mudou de {s0} para {s}.")
        for t in alertas_hoje:
            if t not in prev.get("alertas", []):
                mud.append("Alerta novo: " + t)
        for k, rot in (("selic", "Selic meta"), ("cdi", "CDI")):
            if ok(snap.get(k)) and ok(prev.get(k)) and abs(snap[k] - prev[k]) >= 0.005:
                mud.append(f"{rot} foi de {fbr(prev[k])}% para {fbr(snap[k])}% a.a.")
        if ok(snap.get("dolar")) and ok(prev.get("dolar")) and prev["dolar"]:
            v = snap["dolar"] / prev["dolar"] - 1
            if abs(v) >= 0.01:
                mud.append(f"Dólar {'subiu' if v > 0 else 'caiu'} {fbr(abs(v) * 100, 1)}% (R$ {fbr(prev['dolar'], 4)} → R$ {fbr(snap['dolar'], 4)}).")
    base_txt = (f"Comparado com a leitura de {dt.date.fromisoformat(prev_data).strftime('%d/%m/%Y')}." if prev_data
                else "Primeira leitura guardada: as comparações começam amanhã.")

    # ---------- gravacao nas abas Historico / Snapshots (do Plano_de_Acao.xlsx)
    wh = wb.create_sheet("Historico")
    for j, c in enumerate(HIST_COLS, 1):
        cel = wh.cell(1, j, c)
        cel.font, cel.fill = F_CAB, AZUL
        wh.column_dimensions[get_column_letter(j)].width = 18
    for i, l in enumerate(hist, 2):
        for j, v in enumerate(l, 1):
            cel = wh.cell(i, j, v if not (isinstance(v, float) and math.isnan(v)) else None)
            cel.font = F_N
            if j > 1:
                cel.number_format = N2 if j >= 7 else RS
    wh.freeze_panes = "A2"
    wh.cell(len(hist) + 3, 1, "Uma linha por dia, gravada pelo robô (o último valor do dia vale). Não edite: alimenta o gráfico de desempenho do Dashboard.").font = F_NOTA
    ws_ = wb.create_sheet("Snapshots")
    ws_.cell(1, 1, "Data").font = F_B
    ws_.cell(1, 2, "Foto do dia (JSON, usada para 'o que mudou desde ontem')").font = F_B
    ws_.column_dimensions["B"].width = 80
    for i, d in enumerate(sorted(snaps)[-45:], 2):
        ws_.cell(i, 1, d)
        ws_.cell(i, 2, _j.dumps(snaps[d], ensure_ascii=False, default=float)[:30000])

    # ---------- metas
    meta = {"pat": float(par.get("meta_pat") or 0), "renda": float(par.get("meta_renda") or 0),
            "prazo": float(par.get("meta_prazo") or 0), "ret": float(par.get("meta_ret") or 0)}

    # ---------- proventos
    ativos = [x for x in acoes + fiis if x["qtd"] > 0]
    cal, usou_dy = calendario_proventos(p.get("proventos"), ativos, hoje)
    prov12 = sum(c["total"] for c in cal)
    rf_mes = sum(((x["anos"][0] - x["hoje"]) / 12) for x in linhas_rf) if linhas_rf else 0.0

    # ---------- risco
    inst = {}
    for x in linhas_rf:
        k = x["inst"]
        d = inst.setdefault(k, {"inst": k, "valor": 0.0, "fgc": x["fgc"], "itens": []})
        d["valor"] += x["hoje"]
        d["itens"].append(x["nome"])
        d["fgc"] = d["fgc"] and x["fgc"]
    for x in acoes + fiis:
        inst.setdefault("Corretora (ações e FIIs)", {"inst": "Corretora (ações e FIIs)", "valor": 0.0, "fgc": None, "itens": []})
        inst["Corretora (ações e FIIs)"]["valor"] += x["val"]
    limite = float(par.get("r_fgc") or 250000)
    inst_l = sorted(inst.values(), key=lambda d: -d["valor"])
    idx = {}
    for x in linhas_rf:
        idx[x["idx"]] = idx.get(x["idx"], 0) + x["hoje"]
    if total_ac:
        idx["Ações"] = total_ac
    if total_fi:
        idx["FIIs"] = total_fi
    venc = {"Sem vencimento (liquidez diária)": 0.0, "Até 30 dias": 0.0, "31 a 90 dias": 0.0, "91 dias a 1 ano": 0.0, "Mais de 1 ano": 0.0}
    prox_venc = []
    for x in linhas_rf:
        d = x.get("dias")
        if d is None:
            venc["Sem vencimento (liquidez diária)"] += x["hoje"]
        else:
            k = "Até 30 dias" if d <= 30 else "31 a 90 dias" if d <= 90 else "91 dias a 1 ano" if d <= 365 else "Mais de 1 ano"
            venc[k] += x["hoje"]
            if d <= float(par.get("r_venc") or 60):
                prox_venc.append((x["nome"], x["venc"], d, x["hoje"]))
    venc["Sem vencimento (liquidez diária)"] += total_ac + total_fi
    pesos = sorted([[x["tk"], x["val"] / pat if pat else 0] for x in acoes + fiis], key=lambda t: -t[1])
    risco = {"inst": [{"inst": d["inst"], "valor": d["valor"], "peso": d["valor"] / pat if pat else 0, "fgc": d["fgc"],
                       "excesso": max(0.0, d["valor"] - limite) if d["fgc"] else 0.0, "itens": d["itens"]} for d in inst_l],
             "limite_fgc": limite, "idx": [[k, v] for k, v in sorted(idx.items(), key=lambda t: -t[1])],
             "venc": [[k, v] for k, v in venc.items()], "pesos": pesos[:8], "conc_max": float(par.get("conc_max") or 10),
             "inst_max": float(par.get("r_inst") or 50)}

    # ---------- o que fazer hoje
    hj = []
    if al_df is not None and not al_df.empty:
        for n, t in zip(al_df["Nível"], al_df["Alerta"]):
            if str(n).upper() == "ALTO":
                hj.append({"n": "alta", "t": "Alerta alto", "d": str(t)})
    for s in sinais:
        if s["Sinal"] in ("VENDER", "REDUZIR"):
            hj.append({"n": "alta", "t": f"{s['Sinal'].capitalize()} {s['Ativo']} (sinal pelos seus critérios)", "d": s["Motivo"]})
    for nome, v, d, val in prox_venc:
        hj.append({"n": "alta" if d <= 15 else "media", "t": f"{nome} vence em {d} dia(s)", "d": f"Vencimento {v}, R$ {fbr(val)}. Decida o que fazer com o dinheiro."})
    for d in inst_l:
        if d["fgc"] and d["valor"] > limite:
            hj.append({"n": "alta", "t": f"{d['inst']}: R$ {fbr(d['valor'] - limite)} acima da cobertura do FGC",
                       "d": f"R$ {fbr(d['valor'])} na instituição; o FGC cobre até R$ {fbr(limite, 0)} por CPF e por instituição."})
        elif pat and d["inst"] != "Tesouro Nacional" and d["valor"] / pat * 100 > risco["inst_max"] and len(inst_l) > 1:
            hj.append({"n": "media", "t": f"{fbr(d['valor'] / pat * 100, 0)}% do patrimônio em {d['inst']}",
                       "d": f"Limite de atenção: {risco['inst_max']:g}%. Concentração em uma só instituição é risco de crédito."})
    if al_df is not None and not al_df.empty:
        med = [str(t) for n, t in zip(al_df["Nível"], al_df["Alerta"]) if str(n).upper() != "ALTO"]
        if med:
            hj.append({"n": "media", "t": f"{len(med)} alerta(s) de atenção no painel", "d": " · ".join(med[:3]) + (" …" if len(med) > 3 else "")})
    alvo = ctx.get("alvo") or {}
    atual = ctx.get("atual") or {}
    if alvo and pat:
        faltas = [(c, alvo[c] / 100 * pat - atual.get(c, 0)) for c in ("Renda Fixa", "Ações", "FIIs")]
        grandes = [(c, f) for c, f in faltas if f > 0 and f / pat > 0.02]
        if grandes:
            hj.append({"n": "info", "t": f"Distância até o alvo do perfil {ctx['perfil']}",
                       "d": "; ".join(f"faltam R$ {fbr(f, 0)} em {c}" for c, f in grandes) + ". O próximo aporte pode ir para essas classes."})
    n_ok = len(aprov_ac) + len(aprov_fi)
    if n_ok:
        hj.append({"n": "info", "t": f"{len(aprov_ac)} ação(ões) e {len(aprov_fi)} FII(s) atendem a todos os critérios",
                   "d": "Mais bem colocados: " + ", ".join((aprov_ac[:2] + aprov_fi[:2])) + ". Veja a ficha e confirme antes de decidir."})
    prox = next((c for c in cal[:2] if c["total"] > 0), None)
    if prox:
        hj.append({"n": "info", "t": f"Proventos esperados em {prox['mes']}: R$ {fbr(prox['total'])}",
                   "d": "Estimativa pelos pagamentos dos últimos 12 meses × quantidade atual."})
    return {"meta": meta, "hist": desempenho, "mud": mud, "mud_base": base_txt, "hoje": hj, "prov": cal, "prov12": prov12,
            "prov_dy": usou_dy, "rf_mes": rf_mes, "risco": risco, "aplicado": aplicado, "dolar": m.get("Dólar comercial (R$)"),
            "n_hist": len(hist)}


# ------------------------------------------------------------ dashboard em pagina web
HTML_DASH = r'''<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Dashboard da carteira</title>
<style>
:root{--bg:#F4F6F8;--card:#fff;--bd:#E1E6EB;--tx:#14202B;--mu:#556270;--rf:#0B5D57;--ac:#3B5BDB;--fi:#C96A12;--bar:#EDF0F3}
@media(prefers-color-scheme:dark){:root{--bg:#0F1720;--card:#17212C;--bd:#2A3846;--tx:#E8EDF2;--mu:#9AA9B8;--bar:#26323F}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--tx);font-family:Inter,system-ui,-apple-system,Segoe UI,Arial,sans-serif;font-size:14px}
.w{max-width:1280px;margin:0 auto;padding:28px 24px 48px;display:flex;flex-direction:column;gap:20px}
h1{margin:2px 0 0;font-size:28px}h2{margin:10px 0 0;font-size:19px}h3{margin:0;font-size:15px}
.kick{font-size:12px;font-weight:600;letter-spacing:.08em;color:var(--mu)}
.mu{color:var(--mu);font-size:12px}
.row{display:flex;flex-wrap:wrap;gap:16px}
.card{background:var(--card);border:1px solid var(--bd);border-radius:14px;padding:18px 20px;flex:1 1 220px;min-width:0}
.card .l{font-size:13px;color:var(--mu);display:flex;flex-wrap:wrap;gap:6px;align-items:center}
.card .v{font-size:26px;font-weight:700;margin-top:6px}.card .s{font-size:12px;color:var(--mu);margin-top:4px}
.big{padding:20px 24px}
select{font:inherit;font-weight:600;color:var(--rf);background:rgba(11,93,87,.12);border:1px solid var(--rf);border-radius:7px;padding:3px 8px;cursor:pointer}
table{width:100%;border-collapse:collapse;font-size:13px;margin-top:14px}
th{text-align:left;font-weight:600;color:var(--mu);font-size:12px;padding:8px;border-bottom:1px solid var(--bd)}
td{padding:9px 8px;border-bottom:1px solid var(--bar)}th.r,td.r{text-align:right}tr.t td{font-weight:700;border-bottom:0;border-top:1px solid var(--bd)}
.sc{overflow-x:auto}
.dn{display:flex;flex-wrap:wrap;gap:22px;align-items:center;margin-top:14px}
.lg{display:flex;flex-direction:column;gap:8px;font-size:13px;min-width:200px}.lg i{display:inline-block;width:12px;height:12px;border-radius:3px;margin-right:8px}
.bar{height:12px;background:var(--bar);border-radius:6px;margin-top:5px;position:relative}.bar b{display:block;height:12px;border-radius:6px}
.bh{display:flex;justify-content:space-between;font-size:13px}
.mk{position:absolute;top:-4px;width:3px;height:20px;background:var(--tx);border-radius:2px}
.vazio{font-size:13px;color:var(--mu);margin-top:12px}
.clk{cursor:pointer;border-radius:8px;padding:2px 6px;margin-left:-6px;margin-right:-6px}.clk:hover{background:var(--bar)}
.ov{position:fixed;inset:0;background:rgba(10,18,26,.55);display:flex;align-items:center;justify-content:center;z-index:50;padding:16px}
.mo{background:var(--card);color:var(--tx);border:1px solid var(--bd);border-radius:14px;max-width:760px;width:100%;max-height:90vh;overflow:auto;box-shadow:0 20px 60px rgba(0,0,0,.35)}
.mh{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;padding:18px 22px;border-bottom:1px solid var(--bd);position:sticky;top:0;background:var(--card)}
.mh h3{font-size:20px}.mx{font:inherit;font-size:22px;line-height:1;border:0;background:none;color:var(--mu);cursor:pointer}
.gd{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:10px;padding:18px 22px}
.gd div{border:1px solid var(--bd);border-radius:10px;padding:10px 12px}.gd small{display:block;color:var(--mu);font-size:12px}.gd b{font-size:16px}
.gd .w2{grid-column:1/-1}.tag{display:inline-block;padding:2px 10px;border-radius:99px;font-size:12px;font-weight:700;background:var(--bar)}
.aviso{font-size:12px;color:var(--mu);line-height:1.6}
.tag{font-size:11px;background:rgba(11,93,87,.15);color:var(--rf);padding:1px 6px;border-radius:6px}
</style></head><body><div class="w" id="app"></div>
<script>
const D=__DADOS__;
const HR=D.horizontes;
const $=s=>document.querySelector(s);
const fmt=(v,d=2)=>(v==null||isNaN(v))?'n/d':Number(v).toLocaleString('pt-BR',{minimumFractionDigits:d,maximumFractionDigits:d});
const R=v=>'R$ '+fmt(v);const P=(v,d=1)=>fmt(v*100,d)+'%';
let S={aporte:D.aporte,perfil:D.perfil,rf_card:'2 anos',dv_card:'2 anos',rf_tab:'12 meses',ac_tab:'12 meses',fi_tab:'12 meses',pmt:null};
try{const o=JSON.parse(localStorage.getItem('dash_sel')||'{}');for(const k in o)if(k in S&&(k=='perfil'?D.alvos[o[k]]:(k=='aporte'||k=='pmt')?(typeof o[k]=='number'&&o[k]>=0):HR.includes(o[k])))S[k]=o[k];}catch(e){}
const save=()=>{try{localStorage.setItem('dash_sel',JSON.stringify(S))}catch(e){}};
const hi=r=>HR.indexOf(r);const anos=r=>hi(r)==0?1:hi(r)+1;
const fator=r=>{const n=anos(r),g=D.div_cresc/100;return g==0?n:(Math.pow(1+g,n)-1)/g};
const sel=(k,opts)=>`<select data-k="${k}" aria-label="Escolha">`+opts.map(o=>`<option${o==S[k]?' selected':''}>${o}</option>`).join('')+'</select>';
const donut=(items,cores,centro)=>{const tot=items.reduce((a,x)=>a+x.v,0)||1;let off=0;const r=62,c=2*Math.PI*r;
 let arcs=items.map((x,i)=>{const f=x.v/tot;const a=`<circle cx="80" cy="80" r="${r}" fill="none" stroke="${cores[i%cores.length]}" stroke-width="30" stroke-dasharray="${f*c} ${c}" stroke-dashoffset="${-off*c}" transform="rotate(-90 80 80)"/>`;off+=f;return a}).join('');
 const lg=items.map((x,i)=>`<div><i style="background:${cores[i%cores.length]}"></i>${x.n} <b style="float:right;margin-left:14px">${P(x.v/tot)}</b></div>`).join('');
 return `<div class="dn"><svg width="160" height="160" viewBox="0 0 160 160" role="img" aria-label="Gráfico de rosca">${arcs}<text x="80" y="76" text-anchor="middle" font-size="11" fill="currentColor" opacity=".6">Total</text><text x="80" y="94" text-anchor="middle" font-size="14" font-weight="700" fill="currentColor">${centro}</text></svg><div class="lg">${lg}</div></div>`};
const barras=(lista,cor,f,cls)=>lista.length?lista.map(x=>`<div ${cls&&x[2]?`class="clk" data-f="${cls}:${x[2]}" title="Clique para ver todos os dados"`:''} style="margin-top:12px"><div class="bh"><b>${x[0]}</b><span>${f(x[1])}</span></div><div class="bar"><b style="width:${Math.max(2,Math.min(100,x[1]*100/(Math.max(...lista.map(y=>y[1]))||1)))}%;background:${cor}"></b></div></div>`).join(''):'<div class="vazio">Sem dados nesta atualização.</div>';
const notaPC='<div class="mu" style="margin-top:8px">Preço de compra = preço pago, buscado em Minhas compras e vendas. * Mais de uma compra: média dos preços pagos. Investido = quantidade × preço médio com taxas.</div>';
function tabRV(lst,key,rot){
 if(!lst.length)return `<div class="vazio">Você ainda não tem ${rot} cadastrados na aba Renda Variavel de Minha_Carteira.xlsx. Quando cadastrar, esta área mostra o gráfico e a tabela.</div>`;
 const f=fator(S[key]);const tv=lst.reduce((a,x)=>a+x.val,0);
 const rows=lst.map(x=>`<tr><td>${x.nome}</td><td>${x.tk}</td><td class="r">${fmt(x.qtd,0)}</td><td class="r">${x.pc==null?'n/d':fmt(x.pc)}${x.ncompras>1?' *':''}</td><td class="r">${x.cot==null?'n/d':fmt(x.cot)}</td><td class="r">${x.inv==null?'n/d':fmt(x.inv)}</td><td class="r">${fmt(x.val)}</td><td class="r">${fmt(x.anual/12)}</td><td class="r">${fmt(x.anual*f)}</td></tr>`).join('');
 const sm=k=>lst.reduce((a,x)=>a+(x[k]||0),0);
 return donut(lst.map(x=>({n:x.nome,v:x.val})),key=='ac_tab'?['#3B5BDB','#93A5F0','#5C7CE6','#BAC6F5','#2F46A8']:['#C96A12','#EBB682','#DE8A3A','#F3D2B0','#9A4E0C'],'R$ '+fmt(tv,0))+
 `<div class="mu" style="margin-top:12px">Horizonte: ${sel(key,HR)}</div><div class="sc"><table><thead><tr><th>${key=='ac_tab'?'Empresa':'Fundo'}</th><th>Ticker</th><th class="r">Quantidade</th><th class="r">Preço de compra (R$)</th><th class="r">Cotação atual (R$)</th><th class="r">Investido (R$)</th><th class="r">Valor atual (R$)</th><th class="r">Dividendo próx. mês (R$)</th><th class="r">Dividendo acumulado (R$)</th></tr></thead><tbody>${rows}<tr class="t"><td>Total</td><td></td><td class="r">${fmt(sm('qtd'),0)}</td><td></td><td></td><td class="r">${fmt(sm('inv'))}</td><td class="r">${fmt(sm('val'))}</td><td class="r">${fmt(sm('anual')/12)}</td><td class="r">${fmt(sm('anual')*f)}</td></tr></tbody></table></div>`+notaPC}
const cor=v=>v==null?'':(v<0?'style="color:#C0392B"':'style="color:#1E8449"');
function lotesHTML(){
 if(!D.compras.length)return '<div class="vazio">Nenhuma operação registrada ainda. Lance cada compra na aba Compras de Minha_Carteira.xlsx: o robô calcula o preço médio e mostra o histórico aqui.</div>';
 const rows=D.compras.map(x=>`<tr><td>${x.data}</td><td>${x.nome}</td><td>${x.tk}</td><td>${x.cart}</td><td>${x.tipo}</td><td class="r">${fmt(x.qtd,0)}</td><td class="r">${fmt(x.preco)}</td><td class="r">${fmt(x.taxas)}</td><td class="r">${fmt(x.valor)}</td><td class="r">${x.cot==null?'':fmt(x.cot)}</td><td class="r" ${cor(x.res)}>${x.res==null?'':fmt(x.res)}</td><td class="r" ${cor(x.pct)}>${x.pct==null?'':P(x.pct)}</td></tr>`).join('');
 return `<div class="sc"><table><thead><tr><th>Data</th><th>Empresa / fundo</th><th>Ticker</th><th>Carteira</th><th>Tipo</th><th class="r">Quantidade</th><th class="r">Preço unitário (R$)</th><th class="r">Taxas (R$)</th><th class="r">Valor da operação (R$)</th><th class="r">Cotação atual (R$)</th><th class="r">Resultado (R$)</th><th class="r">Resultado (%)</th></tr></thead><tbody>${rows}</tbody></table></div>`}
const X=D.x||null;
const NIV={alta:['#C0392B','Prioridade'],media:['#D68910','Atenção'],info:['#3B5BDB','Informativo']};
const itens=l=>l.map(x=>`<div style="display:flex;gap:12px;padding:10px 0;border-bottom:1px solid var(--bar)"><span style="flex:0 0 auto;align-self:flex-start;font-size:11px;font-weight:700;color:#fff;background:${NIV[x.n][0]};border-radius:99px;padding:2px 10px">${NIV[x.n][1]}</span><div><b>${x.t}</b><div class="mu" style="margin-top:2px;font-size:13px">${x.d}</div></div></div>`).join('');
function lineSVG(series,labels,fy,h=250){const w=780,pl=56,pr=14,pt=12,pb=26,n=labels.length;
 const all=series.flatMap(s=>s.v.filter(x=>x!=null));if(!all.length)return '';let mn=Math.min(...all),mx=Math.max(...all);if(mx-mn<1e-9){mn-=1;mx+=1}
 const gx=i=>pl+(n<2?0:i*(w-pl-pr)/(n-1)),gy=v=>pt+(h-pt-pb)*(1-(v-mn)/(mx-mn));
 const rg=mx-mn,dc=rg<0.05?3:rg<0.5?2:1;let g='';for(let k=0;k<=4;k++){const v=mn+(mx-mn)*k/4,y=gy(v);g+=`<line x1="${pl}" x2="${w-pr}" y1="${y}" y2="${y}" stroke="currentColor" opacity=".12"/><text x="${pl-6}" y="${y+4}" text-anchor="end" font-size="11" fill="currentColor" opacity=".6">${fy(v,dc)}</text>`}
 const passo=Math.max(1,Math.ceil(n/7));let xl='';for(let i=0;i<n;i+=passo)xl+=`<text x="${gx(i)}" y="${h-6}" text-anchor="${i>=n-passo&&n>7?'end':'middle'}" font-size="11" fill="currentColor" opacity=".6">${labels[i]}</text>`;
 const ln=series.map(s=>{const pts=s.v.map((v,i)=>v==null?null:`${gx(i).toFixed(1)},${gy(v).toFixed(1)}`).filter(Boolean);
  return n==1?`<circle cx="${gx(0)}" cy="${gy(s.v[0])}" r="4" fill="${s.c}"/>`:`<polyline points="${pts.join(' ')}" fill="none" stroke="${s.c}" stroke-width="${s.w||2.2}" ${s.d?'stroke-dasharray="6 4"':''} stroke-linejoin="round"/>`}).join('');
 const lg=series.map(s=>`<span style="margin-right:16px;white-space:nowrap"><i style="display:inline-block;width:14px;height:3px;background:${s.c};margin-right:6px;vertical-align:middle"></i>${s.n}</span>`).join('');
 return `<div class="mu" style="margin-top:10px">${lg}</div><svg viewBox="0 0 ${w} ${h}" style="width:100%;height:auto;margin-top:6px" role="img" aria-label="Gráfico de linhas">${g}${xl}${ln}</svg>`}
function barH(rot,v,max,cor,dir,marca){return `<div style="margin-top:12px"><div class="bh"><b>${rot}</b><span>${dir}</span></div><div class="bar"><b style="width:${Math.max(1,Math.min(100,max?v/max*100:0))}%;background:${cor}"></b>${marca!=null?`<span class="mk" style="left:${Math.min(100,marca)}%"></span>`:''}</div></div>`}
function secHoje(c){if(!X)return '';let l=[...X.hoje];
 if(c.AP>0){const t=c.nomes.map((n,i)=>c.dest[i]>0?`${R(c.dest[i])} em ${n}`:'').filter(Boolean).join(', ');l.unshift({n:'alta',t:`Aporte de ${R(c.AP)} para distribuir`,d:t?`Pelo alvo do perfil ${S.perfil}: ${t}.`:'Todas as classes já estão no alvo; nada a direcionar.'})}
 const mud=X.mud.length?X.mud.map(m=>`<div style="padding:7px 0;border-bottom:1px solid var(--bar)">• ${m}</div>`).join(''):'<div class="vazio">Sem mudanças relevantes desde a última leitura.</div>';
 return `<section class="row"><div class="card big" style="flex:1 1 620px"><h3>O que fazer hoje</h3><div class="mu">Sinais mecânicos pelos seus critérios e limites (aba Parametros), em ordem de prioridade. Não é recomendação de investimento.</div>${l.length?itens(l):'<div class="vazio">Nada urgente hoje. Carteira dentro dos limites que você definiu.</div>'}</div>
 <div class="card big" style="flex:1 1 380px"><h3>O que mudou desde ontem</h3><div class="mu">${X.mud_base}</div><div style="margin-top:8px;font-size:13px">${mud}</div></div></section>`}
let MT={};try{MT=JSON.parse(localStorage.getItem('dash_meta')||'{}')||{}}catch(e){}
const metaEf=()=>{const b=X.meta,o=k=>(typeof MT[k]=='number'&&MT[k]>=0)?MT[k]:b[k];return {pat:o('pat'),renda:o('renda'),prazo:o('prazo'),ret:o('ret')}};
const brIn=v=>Number(v||0).toFixed(2).replace('.',',');
const parseBR=t=>{const s=String(t).trim().replace(/\s/g,'').replace(/\./g,'').replace(',','.');return parseFloat(s)};
const inMeta=(id,v)=>`<input id="mt_${id}" type="text" inputmode="decimal" value="${id=='pat'?brIn(v):v}" style="font:inherit;font-weight:600;width:${id=='pat'?'150':'70'}px;padding:3px 8px;border-radius:7px;border:1px solid var(--rf);background:rgba(11,93,87,.12);color:var(--rf)">`;
function secMeta(c){if(!X)return '';const M=metaEf();
 const i=M.ret>0?Math.pow(1+M.ret/100,1/12)-1:0,n=Math.round(M.prazo*12);
 const fv=c.pat*Math.pow(1+i,n);const pmt=(n>0&&i>0&&M.pat>0)?Math.max(0,(M.pat-fv)*i/(Math.pow(1+i,n)-1)):0;
 const rendaAt=X.rf_mes+X.prov12/12;const pPat=M.pat>0?c.pat/M.pat:0,pRen=M.renda>0?rendaAt/M.renda:0;
 const anosSem=(c.pat>0&&M.pat>c.pat&&M.ret>0)?Math.log(M.pat/c.pat)/Math.log(1+M.ret/100):0;
 const patNec=M.ret>0?M.renda*12/(M.ret/100):null;
 const pm=S.pmt==null?Math.round(pmt/10)*10:S.pmt;
 const ser=[],lab=[],alvo=[];const anosN=Math.max(1,Math.ceil(M.prazo));for(let a=0;a<=anosN;a++){const k=a*12;lab.push(a==0?'hoje':a+(a==1?' ano':' anos'));ser.push(c.pat*Math.pow(1+i,k)+(i>0?pm*(Math.pow(1+i,k)-1)/i:pm*k));alvo.push(M.pat)}
 const fim=ser[ser.length-1];
 return `<section class="card big"><h3>Meta e progresso</h3><div class="mu" style="margin-top:8px;line-height:2.2">Patrimônio-alvo (R$): ${inMeta('pat',M.pat)} &nbsp; Prazo (anos): ${inMeta('prazo',M.prazo)} ${Object.keys(MT).length?'<button id="mt_rst" style="font:inherit;cursor:pointer;margin-left:8px">Voltar aos valores da planilha</button>':''}</div><div class="mu">Digite e tecle Enter: o cálculo muda na hora e fica salvo neste navegador. A planilha (aba Parametros, seção Metas) é o ponto de partida e não muda por aqui.</div>
 <div class="row" style="margin-top:6px">
  <div style="flex:1 1 320px">${barH('Patrimônio',pPat,1,'var(--rf)',`${R(c.pat)} de ${R(M.pat)} · ${P(pPat)}`)}
  <div class="mu" style="margin-top:12px;line-height:1.6">Sem novos aportes, no retorno de ${fmt(M.ret,1)}% ao ano, o patrimônio chegaria à meta em <b>${anosSem>0?fmt(anosSem,1)+' anos':'—'}</b>.</div></div>
  <div style="flex:1 1 320px"><div class="card" style="margin:0"><div class="l">Aporte mensal necessário para chegar a ${R(M.pat)} em ${fmt(M.prazo,0)} anos</div><div class="v">${R(pmt)}</div><div class="s">por mês, a ${fmt(M.ret,1)}% ao ano líquido</div></div>
  <div class="mu" style="margin-top:10px">Simular aporte mensal (R$): <input id="pmt" type="text" inputmode="decimal" value="${brIn(pm)}" style="font:inherit;font-weight:600;width:120px;padding:3px 8px;border-radius:7px;border:1px solid var(--rf);background:rgba(11,93,87,.12);color:var(--rf)"> em ${fmt(M.prazo,0)} anos: <b>${R(fim)}</b> (${fim>=M.pat?'meta atingida':'faltam '+R(M.pat-fim)})</div></div></div>
 ${lineSVG([{n:'Patrimônio projetado com o aporte simulado',c:'#0B5D57',v:ser},{n:'Meta',c:'#C0392B',v:alvo,d:1,w:1.6}],lab,v=>fmt(v/1000,0)+' mil')}</section>`}
function secDesemp(){if(!X)return '';const H=X.hist;
 if(!H)return `<section class="card big"><h3>Desempenho e comparação com o mercado</h3><div class="vazio">O robô começou a guardar o histórico diário da carteira agora. O gráfico (carteira × CDI × IPCA × Ibovespa × IFIX) aparece a partir do segundo dia e fica mais útil a cada semana.</div></section>`;
 const u=a=>a[a.length-1],r=a=>u(a)/100-1;const cdi=r(H.cdi);
 const linhas=[['Sua carteira',H.carteira],['CDI',H.cdi],['IPCA (inflação)',H.ipca],['Ibovespa',H.ibov],['IFIX',H.ifix]].map(x=>`<tr><td>${x[0]}</td><td class="r" ${cor(r(x[1]))}>${P(r(x[1]),2)}</td><td class="r">${cdi>0.00005?fmt(r(x[1])/cdi*100,0)+'% do CDI':'—'}</td></tr>`).join('');
 return `<section class="card big"><h3>Desempenho e comparação com o mercado</h3><div class="mu">Desde ${H.inicio}. A carteira é medida pela rentabilidade, sem contar o dinheiro novo que você aportou. Índices em base 100.</div>
 ${lineSVG([{n:'Sua carteira',c:'#0B5D57',v:H.carteira,w:3},{n:'CDI',c:'#7F8C8D',v:H.cdi},{n:'IPCA',c:'#C96A12',v:H.ipca},{n:'Ibovespa',c:'#3B5BDB',v:H.ibov},{n:'IFIX',c:'#8E44AD',v:H.ifix}],H.datas,(v,d)=>fmt(v,d))}
 <table><thead><tr><th>Comparação no período</th><th class="r">Retorno</th><th class="r">Frente ao CDI</th></tr></thead><tbody>${linhas}</tbody></table></section>`}
function secProv(){if(!X)return '';const C=X.prov,mx=Math.max(...C.map(c=>c.total),1);
 const bars=C.map(c=>`<div style="flex:1;min-width:46px;text-align:center"><div style="height:120px;display:flex;align-items:flex-end;justify-content:center"><div title="${R(c.total)}" style="width:70%;background:var(--ac);border-radius:5px 5px 0 0;height:${Math.max(2,c.total/mx*100)}%"></div></div><div style="font-size:11px;margin-top:4px;font-weight:600">${fmt(c.total,2)}</div><div class="mu" style="font-size:11px">${c.mes}</div></div>`).join('');
 const rows=C.filter(c=>c.total>0).map(c=>`<tr><td>${c.mes}</td><td>${c.itens.map(i=>i[0]+' '+fmt(i[1],2)+(i[2]?' <span class="mu">(dia '+i[2]+')</span>':'')).join(' · ')}</td><td class="r"><b>${fmt(c.total,2)}</b></td></tr>`).join('');
 return `<section class="card big"><h3>Calendário de proventos e renda passiva</h3><div class="mu">Estimativa: o que cada ativo pagou nos últimos 12 meses × sua quantidade atual. A data é a \"data ex\" do ano passado (a partir dela o ativo negocia sem direito ao provento); o dinheiro costuma cair na conta de 1 a 15 dias depois, conforme o comunicado da empresa/fundo. Confirme no comunicado oficial.</div>
 <div class="row" style="margin-top:12px"><div class="card"><div class="l">Proventos em 12 meses</div><div class="v">${R(X.prov12)}</div><div class="s">média de ${R(X.prov12/12)} por mês</div></div>
 <div class="card"><div class="l">Rendimento mensal da renda fixa</div><div class="v">${R(X.rf_mes)}</div><div class="s">líquido de IR, estimativa</div></div>
 <div class="card"><div class="l">Renda mensal total</div><div class="v" style="color:var(--rf)">${R(X.rf_mes+X.prov12/12)}</div><div class="s">renda fixa + proventos</div></div></div>
 <div style="display:flex;gap:4px;margin-top:16px;overflow-x:auto">${bars}</div>
 ${rows?`<div class="sc"><table><thead><tr><th>Mês</th><th>Por ativo (R$) e data estimada</th><th class="r">Total (R$)</th></tr></thead><tbody>${rows}</tbody></table></div>`:'<div class="vazio">Nenhum provento esperado: sem ações/FIIs na carteira ou sem histórico de pagamento.</div>'}
 ${X.prov_dy.length?`<div class="mu" style="margin-top:8px">${X.prov_dy.join(', ')}: sem histórico de pagamentos na fonte; valor estimado pelo dividend yield, dividido igualmente nos meses.</div>`:''}</section>`}
function secRisco(){if(!X)return '';const K=X.risco,L=K.limite_fgc;
 const maxI=Math.max(...K.inst.map(d=>d.valor),L,1);
 const ins=K.inst.map(d=>{const f=d.fgc===null?'sem FGC (ações e FIIs ficam na custódia da B3)':d.fgc?(d.excesso>0?`<span style="color:#C0392B;font-weight:700">R$ ${fmt(d.excesso,0)} acima do FGC</span>`:`dentro do FGC (${P(d.valor/L,0)} do limite)`):'sem cobertura do FGC';
  return barH(d.inst,d.valor,maxI,d.excesso>0?'#C0392B':'var(--rf)',`${R(d.valor)} · ${P(d.peso,d.peso<0.1?1:0)} do patrimônio`,d.fgc?Math.min(100,L/maxI*100):null)+`<div class="mu" style="margin-top:3px">${f}</div>`}).join('');
 const maxX=Math.max(...K.idx.map(x=>x[1]),1),tot=K.idx.reduce((a,x)=>a+x[1],0)||1;
 const idx=K.idx.map(x=>barH(x[0],x[1],maxX,'var(--ac)',`${R(x[1])} · ${P(x[1]/tot,x[1]/tot<0.1?1:0)}`)).join('');
 const mxv=Math.max(...K.venc.map(x=>x[1]),1);const ven=K.venc.map(x=>barH(x[0],x[1],mxv,'var(--fi)',`${R(x[1])}`)).join('');
 const pe=K.pesos.length?K.pesos.map(x=>barH(x[0],x[1],Math.max(K.conc_max/100*1.5,x[1]),x[1]*100>K.conc_max?'#C0392B':'var(--ac)',`${P(x[1],1)} do patrimônio`,K.conc_max/100/Math.max(K.conc_max/100*1.5,x[1])*100)).join(''):'<div class="vazio">Sem ações ou FIIs na carteira.</div>';
 return `<section class="row"><div class="card big" style="flex:1 1 520px"><h3>Concentração por instituição e FGC</h3><div class="mu">Marcador = limite do FGC (${R(L)} por instituição). Se uma instituição quebrar, o FGC devolve até esse valor por CPF.</div>${ins}</div>
 <div class="card big" style="flex:1 1 340px"><h3>Por tipo de rendimento</h3><div class="mu">Dependência de juros (CDI/Selic), inflação (IPCA), prefixado e bolsa</div>${idx}</div>
 <div class="card big" style="flex:1 1 340px"><h3>Quando o dinheiro fica disponível</h3><div class="mu">Renda fixa por vencimento; ações e FIIs vendem em D+2</div>${ven}</div>
 <div class="card big" style="flex:1 1 340px"><h3>Peso de cada ação/FII</h3><div class="mu">Marcador = limite de ${fmt(K.conc_max,0)}% por ativo (aba Parametros)</div>${pe}</div></section>`}
let USD={v:X?X.dolar:null,live:false,pct:null,ts:null};
function updUSD(){const e=document.getElementById('usd');if(!e)return;
 if(USD.v==null){e.innerHTML='';return}
 const hh=USD.ts?USD.ts.toLocaleTimeString('pt-BR'):'';const pat=window.__pat||0;
 e.innerHTML=`<div class="l">Dólar ${USD.live?'<span style="color:#1E8449">● ao vivo</span>':'<span class="mu">(último valor do robô)</span>'}</div><div class="v" style="font-size:22px">R$ ${fmt(USD.v,4)}${USD.pct!=null&&!isNaN(USD.pct)?` <span style="font-size:13px" ${cor(USD.pct)}>${USD.pct>=0?'+':''}${fmt(USD.pct,2)}%</span>`:''}</div><div class="s">${USD.live?`${USD.src} · ${hh}`:''}${pat?` · patrimônio ≈ US$ ${fmt(pat/USD.v,0)}`:''}</div>`}
async function pegaUSD(){
 try{const r=await fetch('https://economia.awesomeapi.com.br/json/last/USD-BRL',{cache:'no-store'});const u=(await r.json()).USDBRL;USD={v:parseFloat(u.bid),pct:parseFloat(u.pctChange),ts:new Date(),live:true,src:'AwesomeAPI'};}
 catch(e){try{const r=await fetch('https://api.frankfurter.dev/v1/latest?base=USD&symbols=BRL');const j=await r.json();USD={v:j.rates.BRL,pct:null,ts:new Date(),live:true,src:'BCE (valor diário)'};}catch(e2){}}
 updUSD()}
function render(){
 const rf=D.rf,ac=D.ac,fi=D.fi,sum=(l,k)=>l.reduce((a,x)=>a+(x[k]||0),0);
 const totRF=sum(rf,'hoje'),totAC=sum(ac,'val'),totFI=sum(fi,'val'),totRV=totAC+totFI,pat=totRF+totRV;
 const um=sum(rf,'um'),a12=rf.reduce((a,x)=>a+x.anos[0],0),aN=rf.reduce((a,x)=>a+x.anos[hi(S.rf_card)],0);
 const dAno=sum(ac,'anual')+sum(fi,'anual'),dMes=dAno/12,dN=dAno*fator(S.dv_card);
 const al=D.alvos[S.perfil]||[0,0,0];const vals=[totRF,totAC,totFI],nomes=['Renda fixa','Ações','FIIs'],cores=['var(--rf)','var(--ac)','var(--fi)'];
 const AP=S.aporte;const base=pat+AP;const falta=al.map((a,i)=>a/100*base-vals[i]);const pos=falta.reduce((a,x)=>a+(x>0?x:0),0);
 const dest=falta.map(x=>(AP>0&&x>0&&pos>0)?AP*x/pos:0);
 window.__pat=pat;const CX={AP,dest,nomes,pat};
 const alocBars=nomes.map((n,i)=>`<div style="margin-top:14px"><div class="bh"><b>${n}</b><span>${P(pat?vals[i]/pat:0)} hoje · alvo ${fmt(al[i],1)}%</span></div><div class="bar" style="height:14px"><b style="width:${Math.min(100,pat?vals[i]/pat*100:0)}%;background:${cores[i]};height:14px"></b><span class="mk" style="left:${al[i]}%"></span></div></div>`).join('');
 const alocTab=nomes.map((n,i)=>`<tr><td>${n}</td><td class="r">${falta[i]>0?R(falta[i]):'acima do alvo'}</td><td class="r"><b>${R(dest[i])}</b></td></tr>`).join('');
 const rfRows=rf.map(x=>`<tr><td>${x.nome}</td><td>${x.tipo}</td><td>${x.idx}</td><td>${x.taxa}</td><td class="r">${fmt(x.hoje)}</td><td class="r">${fmt(x.um)}</td><td class="r">${fmt(x.anos[hi(S.rf_tab)])}</td></tr>`).join('');
 const rfN=rf.reduce((a,x)=>a+x.anos[hi(S.rf_tab)],0);
 const rfBlock=rf.length?donut(rf.map(x=>({n:x.nome,v:x.hoje})),['#0B5D57','#2BA396','#8FD3C9','#5FB9AE','#BFE6E0'],'R$ '+fmt(totRF,0))+
  `<div class="mu" style="margin-top:12px">Horizonte: ${sel('rf_tab',HR)}</div><div class="sc"><table><thead><tr><th>Investimento</th><th>Tipo</th><th>Indexador</th><th>Taxa atual</th><th class="r">Valor hoje (R$)</th><th class="r">Projeção próx. mês (R$)</th><th class="r">Projeção no horizonte (R$)</th></tr></thead><tbody>${rfRows}<tr class="t"><td>Total</td><td></td><td></td><td></td><td class="r">${fmt(totRF)}</td><td class="r">${fmt(um)}</td><td class="r">${fmt(rfN)}</td></tr></tbody></table></div>`
  :'<div class="vazio">Nenhum investimento em Minha_Carteira.xlsx (aba Renda Fixa).</div>';
 $('#app').innerHTML=`
 <header class="row" style="align-items:flex-end;justify-content:space-between"><div><div class="kick">PLANO DE AÇÃO · DASHBOARD</div><h1>Visão geral da carteira</h1>
 <div style="margin-top:12px;display:flex;gap:10px;align-items:center;flex-wrap:wrap"><label class="mu" style="font-size:13px;font-weight:600">Perfil de risco</label>${sel('perfil',Object.keys(D.alvos))}<span class="mu">Aqui vale só para esta tela. O perfil da planilha continua o que está no Excel (${D.perfil}).</span></div></div>
 <div style="display:flex;flex-direction:column;gap:8px;align-items:flex-end"><div class="card" id="usd" style="flex:0 0 auto;min-width:230px"></div><div class="mu">Dados de ${D.gerado}. A página se recarrega sozinha a cada 5 minutos.</div></div></header>
 ${secHoje(CX)}
 <section class="row"><div class="card"><div class="l">Patrimônio total</div><div class="v">${R(pat)}</div><div class="s">${P(pat?totRF/pat:0)} renda fixa · ${P(pat?totRV/pat:0)} renda variável</div></div></section>
 <section class="row">
  <div class="card"><div class="l">Renda fixa</div><div class="v" style="color:var(--rf)">${R(totRF)}</div><div class="s">${rf.length} aplicação(ões) · líquido estimado</div></div>
  <div class="card"><div class="l">Renda fixa: projeção para o próximo mês</div><div class="v">${R(um)}</div><div class="s">+${R(um-totRF)} líquidos de IR (estimativa)</div></div>
  <div class="card"><div class="l">Renda fixa: projeção para daqui a 12 meses</div><div class="v">${R(a12)}</div><div class="s">+${P(totRF?a12/totRF-1:0)} líquido · cenário Focus</div></div>
  <div class="card"><div class="l">Renda fixa: projeção para daqui a ${sel('rf_card',HR)}</div><div class="v">${R(aN)}</div><div class="s">líquido de IR · cenário Focus, sem novos aportes</div></div></section>
 <section class="row">
  <div class="card"><div class="l">Renda variável</div><div class="v" style="color:var(--ac)">${R(totRV)}</div><div class="s">${ac.length+fi.length?ac.length+' ação(ões) · '+fi.length+' FII(s)':'nenhum ativo cadastrado ainda'}</div></div>
  <div class="card"><div class="l">Dividendos estimados: próximo mês</div><div class="v">${R(dMes)}</div><div class="s">média mensal (ações pagam em parcelas irregulares)</div></div>
  <div class="card"><div class="l">Dividendos estimados: próximos 12 meses</div><div class="v">${R(dAno)}</div><div class="s">${P(totRV?dAno/totRV:0)} sobre o valor atual</div></div>
  <div class="card"><div class="l">Dividendos estimados: próximos ${sel('dv_card',HR)}</div><div class="v">${R(dN)}</div><div class="s">soma acumulada; crescimento anual de ${fmt(D.div_cresc,1)}%</div></div></section>
 ${secMeta(CX)}
 <section class="row"><div class="card big" style="flex:1 1 620px"><h3>Alocação atual × alvo do perfil ${S.perfil}</h3><div class="mu">Barra cheia = hoje · marcador vertical = alvo</div>${alocBars}</div>
  <div class="card big" style="flex:1 1 380px"><h3>Destino do aporte de ${R(AP)}</h3><div class="mu" style="margin-top:6px">Aporte (R$): <input id="ap" type="text" inputmode="decimal" value="${brIn(AP)}" style="font:inherit;font-weight:600;width:130px;padding:3px 8px;border-radius:7px;border:1px solid var(--rf);background:rgba(11,93,87,.12);color:var(--rf)"> vale só para esta tela; no Excel (${R(D.aporte)}) mude no Dashboard.</div><table><thead><tr><th>Classe</th><th class="r">Falta para o alvo</th><th class="r">Aporte</th></tr></thead><tbody>${alocTab}</tbody></table></div></section>
 <h2>1 · Minha carteira atual</h2>
 <section class="card big"><h3>Renda fixa</h3><div class="mu">dados reais · valores líquidos de IR, estimados</div>${rfBlock}</section>
 <section class="card big"><h3>Renda variável: ações</h3>${tabRV(ac,'ac_tab','ações')}</section>
 <section class="card big"><h3>Renda variável: fundos imobiliários</h3>${tabRV(fi,'fi_tab','FIIs')}</section>
 <section class="card big"><h3>Minhas compras e vendas</h3><div class="mu">Uma linha por operação, da aba Compras de Minha_Carteira.xlsx</div>${lotesHTML()}</section>
 <h2>2 · Melhores opções agora, pelos seus critérios</h2>
 <section class="row">
  <div class="card big" style="flex:1 1 360px"><h3>Ações · nota de 0 a 100</h3><div class="mu">Os que mais atendem aos critérios da aba Parametros · clique para ver a ficha completa</div>${barras(D.melhores_ac,'var(--ac)',v=>fmt(v*100,1),'ac')}</div>
  <div class="card big" style="flex:1 1 360px"><h3>Fundos imobiliários · nota de 0 a 100</h3><div class="mu">Os que mais atendem aos critérios da aba Parametros · clique para ver a ficha completa</div>${barras(D.melhores_fi,'var(--fi)',v=>fmt(v*100,1),'fi')}</div>
  <div class="card big" style="flex:1 1 360px"><h3>Renda fixa · retorno líquido ao ano</h3><div class="mu">Prazo de 1 a 2 anos, após IR</div>${barras(D.melhores_rf,'var(--rf)',v=>P(v,2))}</div></section>
 <h2>3 · Desempenho e proventos</h2>
 ${secDesemp()}${secProv()}
 <h2>4 · Mapa de risco</h2>
 ${secRisco()}
 <footer class="aviso">${D.aviso}</footer>`;
 updUSD();
 ['pat','prazo'].forEach(k=>{const e=document.getElementById('mt_'+k);if(e)e.addEventListener('change',()=>{const v=parseBR(e.value);if(!isNaN(v)&&v>=0){MT[k]=v;try{localStorage.setItem('dash_meta',JSON.stringify(MT))}catch(x){}render()}})});
 const mr=document.getElementById('mt_rst');if(mr)mr.addEventListener('click',()=>{MT={};try{localStorage.removeItem('dash_meta')}catch(x){}render()});
 const pi=document.getElementById('pmt');if(pi)pi.addEventListener('change',()=>{const v=parseBR(pi.value);S.pmt=(isNaN(v)||v<0)?0:v;save();render()});
 const ap=document.getElementById('ap');if(ap)ap.addEventListener('change',()=>{const v=parseBR(ap.value);S.aporte=(isNaN(v)||v<0)?0:v;save();render()});
 document.querySelectorAll('select[data-k]').forEach(e=>e.addEventListener('change',()=>{S[e.dataset.k]=e.value;save();render()}));
}
function ficha(k){const f=D.fichas[k];if(!f)return;const tk=k.split(':')[1];
 const g=n=>(f.campos.find(c=>c[0]==n)||[0,'n/d'])[1];
 const cls=k.startsWith('ac')?'Ação':'FII';
 const corpo=f.campos.filter(c=>!['Sinal','Ticker','Critérios atendidos','Não atende em'].includes(c[0])).map(c=>`<div><small>${c[0]}</small><b>${c[1]}</b></div>`).join('');
 const ov=document.createElement('div');ov.className='ov';
 ov.innerHTML=`<div class="mo" role="dialog"><div class="mh"><div><h3>${tk}${f.nome?' · '+f.nome:''}</h3><div class="mu" style="margin-top:4px">${cls} · ${f.pos}º no ranking de candidatos · linha da aba Candidatos ${k.startsWith('ac')?'Acoes':'FIIs'} do Plano de Ação</div></div><button class="mx" aria-label="Fechar">×</button></div>
 <div class="gd"><div><small>Sinal</small><b>${g('Sinal')}</b></div><div><small>Critérios atendidos</small><b>${g('Critérios atendidos')}</b></div>${corpo}<div class="w2"><small>Não atende em</small><b style="font-size:14px;font-weight:600">${g('Não atende em')||'-'}</b></div></div>
 <div class="aviso" style="padding:0 22px 18px">Dados mecânicos do painel e dos critérios da aba Parametros. Não é recomendação de compra.</div></div>`;
 const fecha=()=>{ov.remove();document.removeEventListener('keydown',esc);clearTimeout(window.__rl);window.__rl=setTimeout(()=>location.reload(),300000)};
 const esc=e=>{if(e.key=='Escape')fecha()};
 ov.addEventListener('click',e=>{if(e.target==ov||e.target.classList.contains('mx'))fecha()});
 document.addEventListener('keydown',esc);clearTimeout(window.__rl);document.body.appendChild(ov)}
document.addEventListener('click',e=>{const t=e.target.closest&&e.target.closest('.clk');if(t)ficha(t.dataset.f)});
render();
if(!/[?&]pdf=1/.test(location.search)){pegaUSD();setInterval(pegaUSD,30000)}
if(/[?&]pdf=1/.test(location.search)){
 /* modo PDF: uma unica pagina continua, sem cortes; sem recarregar */
 document.body.style.background='#fff';
 const h=Math.ceil(Math.max(document.documentElement.scrollHeight,document.body.scrollHeight))+24;
 const st=document.createElement('style');st.textContent='@page{size:1280px '+h+'px;margin:0}html,body{background:#fff!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}.card,table,tr,.row{break-inside:avoid}';
 document.head.appendChild(st);
}else{window.__rl=setTimeout(()=>location.reload(),300000)}
</script></body></html>
'''


def _navegador():
    """Caminho do Edge ou Chrome (o Windows ja traz o Edge)."""
    import shutil
    cands = [os.environ.get(v, "") + c for v, c in (
        ("ProgramFiles(x86)", r"\Microsoft\Edge\Application\msedge.exe"),
        ("ProgramFiles", r"\Microsoft\Edge\Application\msedge.exe"),
        ("ProgramFiles", r"\Google\Chrome\Application\chrome.exe"),
        ("ProgramFiles(x86)", r"\Google\Chrome\Application\chrome.exe"),
        ("LOCALAPPDATA", r"\Google\Chrome\Application\chrome.exe"))]
    for c in cands:
        if c and os.path.exists(c):
            return c
    return shutil.which("msedge") or shutil.which("google-chrome") or shutil.which("chromium") or shutil.which("chromium-browser")


def gerar_pdf(html, pasta):
    """Grava Dashboard.pdf: uma pagina continua e larga (nada cortado), igual ao Dashboard.html.
    Usa o Edge/Chrome do proprio computador, sem instalar nada. So regrava se os dados mudaram."""
    import hashlib
    import subprocess
    import tempfile
    try:
        nav = _navegador()
        if not html:
            return None
        if not nav:
            log.warning("Dashboard.pdf nao gerado: Edge/Chrome nao encontrado neste computador")
            return None
        destino = Path(pasta) / "Dashboard.pdf"
        txt = Path(html).read_text(encoding="utf-8")
        chave = hashlib.md5(re.sub(r'"gerado": "[^"]*"', "", txt).encode("utf-8")).hexdigest()
        marca = BASE / "cache" / "pdf_hash.txt"
        if destino.exists() and marca.exists() and marca.read_text().strip() == chave:
            return destino
        tmp = Path(pasta) / "~tmp_Dashboard.pdf"
        if tmp.exists():
            tmp.unlink()
        perfil = tempfile.mkdtemp(prefix="dash_pdf_")
        url = Path(html).resolve().as_uri() + "?pdf=1"
        cmd = [nav, "--headless=new", "--disable-gpu", "--no-sandbox", f"--user-data-dir={perfil}",
               "--no-pdf-header-footer", "--window-size=1280,900", "--virtual-time-budget=4000", f"--print-to-pdf={tmp}", url]
        subprocess.run(cmd, timeout=90, capture_output=True)
        try:
            import shutil
            shutil.rmtree(perfil, ignore_errors=True)
        except Exception:  # noqa
            pass
        if not tmp.exists() or tmp.stat().st_size < 5000:
            log.warning("Dashboard.pdf nao gerado (navegador sem resposta)")
            return None
        os.replace(tmp, destino)
        marca.parent.mkdir(exist_ok=True)
        marca.write_text(chave)
        log.info("Dashboard.pdf gravado em %s", destino)
        return destino
    except PermissionError:
        log.warning("Dashboard.pdf aberto em outro programa; fica para a proxima atualizacao")
    except Exception as e:  # noqa
        log.warning("Dashboard.pdf nao gerado: %s", e)
    return None


def gerar_html(dados, pasta):
    """Grava Dashboard.html (pagina unica, abre no navegador, sem internet)."""
    import json
    try:
        txt = HTML_DASH.replace("__DADOS__", json.dumps(dados, ensure_ascii=False, default=float))
        destino = Path(pasta) / "Dashboard.html"
        tmp = destino.with_name("~tmp_Dashboard.html")
        tmp.write_text(txt, encoding="utf-8")
        os.replace(tmp, destino)
        return destino
    except Exception:  # noqa
        return None


# --------------------------------------------------------------- principal
def gerar(pasta, arq_painel=None):
    pasta = Path(pasta)
    arq_painel = Path(arq_painel) if arq_painel else pasta / "Painel_Investimentos.xlsx"
    destino = pasta / ARQ_PLANO
    par = ler_parametros(destino)
    dash_in = ler_dash(destino)
    p = ler_painel(arq_painel)
    m = p["macro"]

    # ---- cenario
    selic, cdi = m.get("Selic meta (% a.a.)"), m.get("CDI (% a.a.)")
    prox = str(AGORA.year + 1)
    fx = {}
    if not p["focus"].empty and prox in p["focus"].columns:
        fx = {str(i): v for i, v in zip(p["focus"]["Indicador"], p["focus"][prox]) if ok(v)}
    selic_fut, ipca_fut = fx.get("Selic"), fx.get("IPCA")
    ipca_mais = None
    t = p["tesouro"]
    if not t.empty:
        ref = t[(t["Título"] == "Tesouro IPCA+") & (t["Prazo (anos)"] >= 4) & (t["Prazo (anos)"] <= 15)]
        if len(ref):
            ipca_mais = float(ref["Taxa compra (% a.a.)"].median())
    juro_real_focus = ((1 + (selic_fut or selic) / 100) / (1 + ipca_fut / 100) - 1) if (selic or selic_fut) and ipca_fut else None
    juro_real = ipca_mais if ipca_mais is not None else juro_real_focus
    if ipca_mais is None:
        ipca_mais = juro_real_focus

    # ---- candidatos
    ca = candidatos(p["acoes"], avaliar_acao, par, juro_real)
    cf = candidatos(p["fiis"], avaliar_fii, par, ipca_mais)
    if not ca.empty:
        ca["Lucro ÷ preço"] = ca["P/L"].map(lambda x: 1 / x if ok(x) and x > 0 else None)
        ca["Margem vs. Graham"] = [(g / c - 1) if ok(g) and ok(c) and c else None
                                   for g, c in zip(ca["Preço justo Graham"], ca["Cotação"])]
    for _df in (ca, cf):
        if not _df.empty:
            _df["Rendimento 12m (R$/cota)"] = [float(d) * float(c) if ok(d) and ok(c) else None
                                                for d, c in zip(_df["DY 12m"], _df["Cotação"])]
    if not cf.empty and ipca_mais is not None:
        cf["Prêmio sobre IPCA+ (p.p.)"] = cf["DY 12m"].map(lambda x: (x - ipca_mais) * 100 if ok(x) else None)

    # ---- alocacao
    total = m.get("Patrimônio total", 0) or 0
    atual = {"Renda Fixa": 0.0, "Ações": 0.0, "FIIs": 0.0}
    if not p["aloc"].empty:
        for c, v in zip(p["aloc"]["Classe"], p["aloc"]["Valor (R$)"]):
            if c in atual and ok(v):
                atual[c] = float(v)
    med_a = ca["Lucro ÷ preço"].dropna().median() if not ca.empty else None
    med_f = cf["DY 12m"].where((cf["DY 12m"] > 0) & (cf["DY 12m"] < 0.25)).median() if not cf.empty else None
    prem_a = float(med_a) - juro_real if ok(med_a) and juro_real is not None else None
    prem_f = float(med_f) - ipca_mais if ok(med_f) and ipca_mais is not None else None
    _, t_prem, t_juros, juros_txt, prem_med, frac_ac = alvo_automatico(selic, selic_fut, prem_a, prem_f)
    perfil = str(par.get("perfil") or "Conservador").strip().capitalize()
    tab_alvos = par["tabela_alvos"]
    if perfil not in tab_alvos:
        perfil = "Conservador"
    alvo = dict(zip(("Renda Fixa", "Ações", "FIIs"), tab_alvos[perfil]))
    tem_alvo = all(ok(v) for v in alvo.values()) and abs(sum(alvo.values()) - 100) < 0.5
    aporte = float(par["aporte"] or 0)
    base = total + aporte

    # ---- sinais da carteira de renda variavel
    sinais = []
    rv = p["rv"]
    n_max = int(par["max_ativos"] or 5)
    for _, h in rv.iterrows() if not rv.empty else []:
        tk, classe = h["Ticker"], h.get("Classe")
        tab = cf if classe == "FIIs" else ca
        lin = tab[tab["Ticker"] == tk] if not tab.empty else pd.DataFrame()
        peso = h.get("Peso na carteira")
        if lin.empty:
            sinais.append({"Sinal": "MANTER", "Ativo": tk, "Classe": classe, "Cotação": h.get("Cotação"),
                           "Peso na carteira": peso, "Score (0-100)": h.get("Score (0-100)"),
                           "Motivo": "Sem indicadores no painel para avaliar; confira o ticker."})
            continue
        l = lin.iloc[0]
        g = gatilhos_venda(l, classe, par)
        abaixo_alvo = (not tem_alvo) or (alvo[classe] / 100 * base - atual.get(classe, 0) > 0)
        if g:
            sinal, motivo = "VENDER", "Gatilho de venda: " + "; ".join(g) + "."
        elif ok(peso) and peso * 100 > par["conc_max"]:
            sinal = "REDUZIR"
            motivo = f"Posição de {fpc(peso)} da carteira, acima do limite de {par['conc_max']:g}%."
        elif l["_todos"] and abaixo_alvo:
            sinal, motivo = "COMPRAR MAIS", "Atende a todos os critérios de compra e há espaço na classe e na posição."
        else:
            sinal = "MANTER"
            motivo = ("Atende aos critérios, mas a classe já está no alvo." if l["_todos"]
                      else f"Sem gatilho de venda. Não atende à compra em: {l['Não atende em']}.")
        sinais.append({"Sinal": sinal, "Ativo": tk, "Classe": classe, "Cotação": h.get("Cotação"),
                       "Peso na carteira": peso, "Score (0-100)": l.get("Score (0-100)"), "Motivo": motivo})

    # ---- plano de compras (valores calculados por formula na planilha, a partir do aporte)
    compras = []
    na_carteira = set(rv["Ticker"]) if not rv.empty else set()
    for classe, tab in (("Ações", ca), ("FIIs", cf)):
        if tab.empty:
            continue
        aptos = tab[tab["_todos"]].head(10)
        for pos, (_, l) in enumerate(aptos.iterrows(), 1):
            compras.append({"Sinal": "COMPRAR MAIS" if l["Ticker"] in na_carteira else "COMPRAR",
                            "Ativo": l["Ticker"], "Classe": classe, "Cotação": l["Cotação"],
                            "Score (0-100)": l["Score (0-100)"], "DY 12m": l["DY 12m"],
                            "Posição na classe": pos, "No plano?": None, "Qtd. de referência": None,
                            "Valor de referência (R$)": None, "_n": len(aptos),
                            "Motivo": "Atende a todos os critérios de compra da aba Parametros."})

    # =============================================================== planilha
    wb = Workbook()
    ws = wb.active
    ws.title = "Plano de Acao"
    titulo(ws, "PLANO DE AÇÃO DA CARTEIRA", 9, f"Gerado em {AGORA:%d/%m/%Y às %H:%M} a partir do Painel ({p['atualizado']})")
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 16
    ws["A3"] = AVISO
    ws["A3"].font = F_NOTA
    r = 5
    ws.cell(r, 1, "1. CENÁRIO DE MERCADO").font = F_SEC
    r += 1
    if selic is not None:
        r = par_linha(ws, r, "Selic hoje (% a.a.)", selic, N2)
    if selic_fut is not None:
        dif = selic_fut - (selic or selic_fut)
        r = par_linha(ws, r, f"Selic esperada no fim de {prox} (Focus)", selic_fut, N2,
                      f"mercado espera {'queda' if dif < 0 else 'alta'} de {fbr(abs(dif))} p.p." if abs(dif) >= 0.01 else "mercado espera estabilidade")
    if ipca_fut is not None:
        r = par_linha(ws, r, f"IPCA esperado em {prox} (Focus, %)", ipca_fut, N2)
    if juro_real_focus is not None:
        r = par_linha(ws, r, "Juro real esperado (Selic ÷ IPCA do Focus)", juro_real_focus, PC)
    if ipca_mais is not None:
        r = par_linha(ws, r, "Tesouro IPCA+ de referência (4 a 15 anos)", ipca_mais, PC,
                      "taxa real sem risco de crédito; é a régua dos critérios de prêmio")
    if not ca.empty:
        med = ca["Lucro ÷ preço"].dropna().median()
        if ok(med):
            r = par_linha(ws, r, "Lucro ÷ preço mediano das ações acompanhadas", float(med), PC,
                          f"prêmio sobre o Tesouro IPCA+: {fbr((med - (juro_real or 0)) * 100, 1)} p.p." if juro_real is not None else None)
    if not cf.empty:
        med = cf["DY 12m"].where((cf["DY 12m"] > 0) & (cf["DY 12m"] < 0.25)).median()
        if ok(med):
            r = par_linha(ws, r, "Dividend yield mediano dos FIIs acompanhados", float(med), PC,
                          f"prêmio sobre o Tesouro IPCA+: {fbr((med - (ipca_mais or 0)) * 100, 1)} p.p." if ipca_mais is not None else None)
    if selic is not None and selic_fut is not None:
        if selic_fut <= selic - 0.5:
            txt = ("Leitura: o mercado projeta juros em queda. Em geral, esse cenário reduz aos poucos o rendimento dos "
                   "pós-fixados (CDI/Selic), valoriza prefixados e IPCA+ já contratados e tende a favorecer ações e FIIs. "
                   "Projeções mudam e podem não se confirmar.")
        elif selic_fut >= selic + 0.5:
            txt = ("Leitura: o mercado projeta juros em alta. Em geral, esse cenário favorece pós-fixados e pressiona "
                   "prefixados, IPCA+ longos, ações e FIIs. Projeções mudam e podem não se confirmar.")
        else:
            txt = "Leitura: o mercado projeta juros estáveis; o rendimento dos pós-fixados tende a se manter no nível atual."
        ws.cell(r, 1, txt).font = F_NOTA
        r += 1
    r += 1
    ws.cell(r, 1, "2. ALOCAÇÃO: ONDE VOCÊ ESTÁ E ONDE QUER CHEGAR").font = F_SEC
    r += 1
    pm = mapa_param()
    P = lambda k: f"Parametros!$B${pm[k]}"
    r_tot = r
    r = par_linha(ws, r, "Patrimônio total (R$)", total, RS)
    r_apo = r
    r = par_linha(ws, r, "Aporte disponível (R$)", f"={P('aporte')}", RS,
                  "digite no Dashboard ou na aba Parametros: esta seção se recalcula na hora, no Excel")
    r = par_linha(ws, r, "Perfil de risco", f"={P('perfil')}", None,
                  "escolha na aba Parametros (Conservador, Moderado, Arrojado ou Personalizado)")
    h = r + 1
    cab = ["Classe", "Valor atual (R$)", "Atual", "Alvo do perfil", "Diferença para o alvo (R$)", "Destino do aporte (R$)"]
    for j, c in enumerate(cab, 1):
        cel = ws.cell(h, j, c)
        cel.font, cel.fill = F_CAB, AZUL
        cel.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        if j > 3 and ws.column_dimensions[get_column_letter(j)].width < 16:
            ws.column_dimensions[get_column_letter(j)].width = 16
    ws.row_dimensions[h].height = 30
    classes = ["Renda Fixa", "Ações", "FIIs"]
    chave_alvo = {"Renda Fixa": "alvo_rf", "Ações": "alvo_acoes", "FIIs": "alvo_fiis"}
    g0, g1 = h + 1, h + 3
    for i, c in enumerate(classes):
        rr = h + 1 + i
        ws.cell(rr, 1, c)
        ws.cell(rr, 2, atual[c]).number_format = RS
        ws.cell(rr, 3, f"=IF($B${r_tot}>0,B{rr}/$B${r_tot},0)").number_format = PC
        ws.cell(rr, 4, f"=N({P(chave_alvo[c])})/100").number_format = PC
        ws.cell(rr, 5, f"=D{rr}*($B${r_tot}+$B${r_apo})-B{rr}").number_format = RS
        ws.cell(rr, 6, f'=IF(AND($B${r_apo}>0,E{rr}>0,SUMIF($E${g0}:$E${g1},">0")>0),'
                       f'$B${r_apo}*E{rr}/SUMIF($E${g0}:$E${g1},">0"),0)').number_format = RS
        for j in range(1, 7):
            ws.cell(rr, j).font, ws.cell(rr, j).border = (F_B if j in (1, 4, 6) else F_N), BORDA
    rc = h + 4
    c = ws.cell(rc, 1, f'=IF(ABS(SUM(D{g0}:D{g1})-1)>0.005,"ATENÇÃO: os três alvos do perfil somam "&ROUND(SUM(D{g0}:D{g1})*100,1)'
                       f'&"% e precisam somar 100%. Corrija a tabela de alvos na aba Parametros.","Alvos do perfil somam 100%.")')
    c.font, c.fill = F_B, AMAR
    ws.cell(rc + 1, 1, "Diferença positiva = falta nessa classe para chegar ao alvo; negativa = sobra. O aporte vai para as "
                       "classes com diferença positiva, na proporção do que falta.").font = F_NOTA
    mk = [f"Leitura do mercado hoje (informativa; não altera seus alvos): prêmio de ações e FIIs sobre o Tesouro IPCA+ "
          + (f"de {fbr(prem_med, 1)} p.p. (referência 4,0), " if prem_med is not None else "n/d, ")
          + f"{juros_txt}. Esses sinais costumam favorecer um pouco mais de renda variável do que a base do perfil "
            f"({t_prem + t_juros:+.1f} p.p.).".replace(".", ",") if (t_prem + t_juros) > 0 else
          f"Leitura do mercado hoje (informativa; não altera seus alvos): {juros_txt}.",
          "Os percentuais de cada perfil estão na tabela no fim da aba Parametros e podem ser editados. "
          "Os alvos são decisão sua: a regra não conhece seus objetivos, prazos nem sua reserva de emergência."]
    for i, t in enumerate(mk):
        ws.cell(rc + 2 + i, 1, t).font = F_NOTA
    r = rc + 2 + len(mk) + 1
    DEST = {"Ações": f"$F${h + 2}", "FIIs": f"$F${h + 3}"}

    ws.cell(r, 1, "3. O QUE FAZER COM O QUE JÁ ESTÁ NA CARTEIRA").font = F_SEC
    r += 1
    if sinais:
        ordem = {"VENDER": 0, "REDUZIR": 1, "COMPRAR MAIS": 2, "MANTER": 3}
        ds = pd.DataFrame(sinais).sort_values("Sinal", key=lambda s: s.map(ordem))
        r = escrever(ws, r, ds, {"Cotação": RS, "Peso na carteira": PC, "Score (0-100)": "0.0"},
                     {"Motivo": 70, "Sinal": 46}, cor_col="Sinal")
    else:
        ws.cell(r, 1, "Você não tem ações nem FIIs cadastrados; não há sinais de venda ou manutenção a calcular.").font = F_N
        r += 2
    if not p["rf"].empty:
        rf = p["rf"].copy()
        linhas = []
        for _, l in rf.iterrows():
            rem = str(l.get("Remuneração") or "")
            mm = re.match(r"([\d.,]+)% do CDI", rem)
            pct = float(mm.group(1).replace(".", "").replace(",", ".")) if mm else None
            if pct and pct <= 3:
                pct *= 100  # painel antigo gravava 100% como 1
                rem = f"{pct:g}% do CDI".replace(".", ",")
            bruto = cdi / 100 * pct / 100 if pct and cdi else None
            futuro = (selic_fut - 0.10) / 100 * pct / 100 if pct and selic_fut else None
            linhas.append({"Investimento": l["Investimento"], "Remuneração": rem,
                           "Valor (R$)": l.get("Valor bruto estimado (R$)"),
                           "Rende hoje (% a.a. bruto)": bruto,
                           f"Renderia com a Selic esperada p/ {prox}": futuro,
                           "Tesouro IPCA+ de referência": ipca_mais})
        ws.cell(r, 1, "Renda fixa atual frente ao cenário").font = F_B
        r = escrever(ws, r + 1, pd.DataFrame(linhas),
                     {"Valor (R$)": RS, "Rende hoje (% a.a. bruto)": PC, f"Renderia com a Selic esperada p/ {prox}": PC,
                      "Tesouro IPCA+ de referência": PC}, {"Investimento": 46, "Remuneração": 18})
        ws.cell(r - 1, 1, "Pós-fixado acompanha a Selic: se ela cair como o Focus projeta, o rendimento cai junto. "
                          "O Tesouro IPCA+ é taxa real (acima da inflação) e oscila de preço antes do vencimento.").font = F_NOTA
        r += 1
    ws.cell(r, 1, "4. CANDIDATOS A COMPRA PELOS SEUS CRITÉRIOS").font = F_SEC
    r += 1
    if compras:
        dc = pd.DataFrame(compras)
        ns = list(dc["_n"])
        dc = dc.drop(columns=["_n"])[["Sinal", "Ativo", "Classe", "Cotação", "Score (0-100)", "DY 12m", "Posição na classe",
                                      "No plano?", "Qtd. de referência", "Valor de referência (R$)", "Motivo"]]
        r0 = r
        r = escrever(ws, r, dc, {"Cotação": RS, "Score (0-100)": "0.0", "DY 12m": PC, "Qtd. de referência": N0,
                                 "Valor de referência (R$)": RS}, {"Motivo": 60, "Sinal": 46}, cor_col="Sinal")
        mx = f"MAX(1,MIN({P('max_ativos')},{{n}}))"
        for i, (cl, n) in enumerate(zip(dc["Classe"], ns), 1):
            rr = r0 + i
            m_ = mx.format(n=n)
            ws.cell(rr, 8, f'=IF(G{rr}<={m_},"SIM","reserva")')
            ws.cell(rr, 9, f'=IF(AND(H{rr}="SIM",{DEST[cl]}>0,D{rr}>0),INT({DEST[cl]}/{m_}/D{rr}),"")').number_format = N0
            ws.cell(rr, 10, f'=IF(I{rr}="","",I{rr}*D{rr})').number_format = RS
            for j in (8, 9, 10):
                ws.cell(rr, j).font, ws.cell(rr, j).border = F_N, BORDA
        ws.cell(r - 1, 1, "Quantidade de referência = verba da classe (Destino do aporte) ÷ nº de ativos no plano ÷ cotação. "
                          "Muda na hora ao digitar o aporte ou o nº máximo de ativos; 'reserva' = candidato além do limite.").font = F_NOTA
        r += 1
    else:
        ws.cell(r, 1, "Nenhum ativo atende a todos os critérios de compra neste momento. Veja nas abas de candidatos "
                      "os que chegaram mais perto, ou ajuste os critérios na aba Parametros.").font = F_N
        r += 2
    ws.cell(r, 1, f"Ações que atendem a todos os critérios: {int(ca['_todos'].sum()) if not ca.empty else 0} de {len(ca)}.   "
                  f"FIIs: {int(cf['_todos'].sum()) if not cf.empty else 0} de {len(cf)}.").font = F_NOTA

    # ---- abas de candidatos
    def aba(nome, tit, df, cols, fmts, regua):
        w = wb.create_sheet(nome)
        titulo(w, tit, len(cols) + 4, regua)
        if df.empty:
            w["A4"] = "Dados indisponíveis no painel."
            return
        d = df.head(40)[["Sinal", "Ticker", "Critérios atendidos"] + cols + ["Não atende em"]]
        escrever(w, 4, d, fmts, {"Não atende em": 60, "Sinal": 14, "Critérios atendidos": 14}, cor_col="Sinal")
        w.freeze_panes = "C5"
        from openpyxl.worksheet.hyperlink import Hyperlink
        v_ = w.cell(3, 1, "◄ Voltar ao Dashboard")
        v_.hyperlink = Hyperlink(ref="A3", location="'Dashboard'!A1", display="Voltar ao Dashboard")
        v_.font = Font(name="Arial", size=10, color="0563C1", underline="single", bold=True)

    aba("Candidatos Acoes", "AÇÕES — MAIS PRÓXIMAS DOS SEUS CRITÉRIOS (40 primeiras)", ca,
        COLS_AC,
        {"Cotação": RS, "Score (0-100)": "0.0", "P/L": N2, "Lucro ÷ preço": PC, "P/VP": N2, "DY 12m": PC, "Rendimento 12m (R$/cota)": RS, "ROE": PC, "Margem líq.": PC,
         "Dív.líq/Patrim.": N2, "Dív.líq/EBITDA": N2, "EBITDA (R$ mi)": N0, "Dív. líquida (R$ mi)": N0, "Preço justo Graham": RS, "Margem vs. Graham": PC, "Preço-teto Bazin": RS,
         "Liquidez diária (R$)": N0},
        f"Régua de mercado: lucro ÷ preço precisa superar {fpc(juro_real, 2) if juro_real is not None else 'n/d'} "
        f"(Tesouro IPCA+) mais {par['a_premio']:g} p.p. de prêmio.")
    aba("Candidatos FIIs", "FIIs — MAIS PRÓXIMOS DOS SEUS CRITÉRIOS (40 primeiros)", cf,
        cols_fi(cf),
        {"Cotação": RS, "Score (0-100)": "0.0", "DY 12m": PC, "Rendimento 12m (R$/cota)": RS, "Prêmio sobre IPCA+ (p.p.)": "0.0", "P/VP": N2,
         "Vacância": PC, "Imóveis": N0, "Liquidez diária (R$)": N0},
        f"Régua de mercado: DY precisa superar {fpc(ipca_mais, 2) if ipca_mais is not None else 'n/d'} "
        f"(Tesouro IPCA+) mais {par['f_premio']:g} p.p. de prêmio.")

    # ---- parametros
    w = wb.create_sheet("Parametros")
    titulo(w, "PARÂMETROS — SEUS CRITÉRIOS", 3,
           "Altere só as células amarelas. Aporte, perfil de risco (que define os alvos) e nº de ativos recalculam na hora, ao digitar no Excel. "
           "Os critérios de compra e venda (liquidez, P/L etc.) valem a partir da próxima atualização do robô (até 15 min). Salve e feche.")
    w.column_dimensions["A"].width = 58
    w.column_dimensions["B"].width = 16
    w.column_dimensions["C"].width = 100
    r = 4
    t_hdr, t_ini = tabela_pos()
    col_alvo = {"alvo_rf": "B", "alvo_acoes": "C", "alvo_fiis": "D"}
    for k, rot, _, exp in PARAMETROS:
        if k.startswith("sec"):
            r += 1
            w.cell(r, 1, rot).font = F_SEC
        else:
            w.cell(r, 1, rot).font = F_N
            if k == "perfil":
                c = w.cell(r, 2, f"=Dashboard!${DASH_PERFIL[0]}${DASH_PERFIL[1:]}")
                c.font, c.number_format = F_B, "General"
            elif k == "aporte":
                c = w.cell(r, 2, f"=Dashboard!${DASH_APORTE[0]}${DASH_APORTE[1:]}")
                c.font, c.number_format = F_B, RS
            elif k in col_alvo:
                c = w.cell(r, 2, f"=IFERROR(INDEX(${col_alvo[k]}${t_ini}:${col_alvo[k]}${t_ini + 3},"
                                 f"MATCH($B${pm['perfil']},$A${t_ini}:$A${t_ini + 3},0)),0)")
                c.font, c.number_format = F_B, "0.0"
            else:
                c = w.cell(r, 2, par[k])
                c.font, c.fill = F_ENT, ENTRADA
                if isinstance(par[k], (int, float)) and abs(par[k]) >= 1000:
                    c.number_format = N0
            if exp:
                w.cell(r, 3, exp).font = F_NOTA
        r += 1
    w.cell(t_hdr - 1, 1, "TABELA DE ALVOS POR PERFIL (% da carteira; cada linha deve somar 100)").font = F_SEC
    for j, cb in enumerate(["Perfil", "Renda Fixa (%)", "Ações (%)", "FIIs (%)", "Soma"], 1):
        cel = w.cell(t_hdr, j, cb)
        cel.font, cel.fill = F_CAB, AZUL
        cel.alignment = Alignment(horizontal="center")
    for i, nome in enumerate(ALVOS_PERFIL):
        rr = t_ini + i
        w.cell(rr, 1, nome).font = F_B
        for j, v in enumerate(par["tabela_alvos"][nome], 2):
            cel = w.cell(rr, j, v)
            cel.font, cel.fill, cel.number_format = F_ENT, ENTRADA, "0.0"
        w.cell(rr, 5, f'=IF(ABS(SUM(B{rr}:D{rr})-100)>0.05,"≠ 100","OK")').font = F_N
    w.cell(t_ini + 4, 1, "Conservador, Moderado e Arrojado seguem a orientação de alocação informada por você; "
                         "Personalizado é livre para os seus próprios percentuais.").font = F_NOTA
    w.cell(t_ini + 6, 1, AVISO).font = F_NOTA
    dv2 = DataValidation(type="list", formula1='"Conservador,Moderado,Arrojado,Personalizado"', allow_blank=False)
    w.add_data_validation(dv2)
    dv2.add(f"B{pm['perfil']}")
    dados_html = construir_dashboard(wb, {"p": p, "par": par, "pm": pm, "dash_in": dash_in, "perfil": perfil, "h_aloc": h,
                             "ca": ca, "cf": cf, "cdi": cdi, "ipca_mais": ipca_mais, "ipca_fut": ipca_fut,
                             "sinais": sinais, "destino": destino, "alvo": alvo, "atual": atual})
    wb.active = 0
    for i, w_ in enumerate(wb.worksheets):
        w_.sheet_view.tabSelected = (i == 0)
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    saida = salvar(wb, destino)
    gerar_pdf(gerar_html(dados_html, pasta), pasta)
    return saida


if __name__ == "__main__":
    pasta_cli = sys.argv[1] if len(sys.argv) > 1 else "."
    print("Plano gravado em", gerar(pasta_cli))
