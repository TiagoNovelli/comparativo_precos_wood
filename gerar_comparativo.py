#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera os relatorios HTML "Comparativo de Precos Wood" a partir das planilhas:
  - Novos precos.xlsx                                  (precos novos, fonte unica)
  - 1 - TABELA DE PRECOS WOOD - CENTURY 28.04.xlsx     (precos antigos, marca Century)
  - 1 - TABELA DE PRECOS WOOD - PV 28.04.xlsx          (precos antigos, marca Pv)
  - 1 - TABELA DE PRECOS WOOD - PRIVATE LABEL 28.04.xlsx (precos antigos, marca Private Label)

Regra de correspondencia (fixa, deterministica):
  Para cada produto novo (MODELO ... + categoria), procura nas planilhas antigas da
  MESMA categoria um produto cujo nome compartilhe a "palavra-chave" do modelo;
  entre todos os pares (variante nova x produto antigo candidato), escolhe o par
  com a MENOR diferenca de dimensoes (C+L+A). Se nao houver candidato na categoria
  (ou a categoria nao existir nas planilhas antigas), o produto cai em
  "Sem Correspondente".

Limitacoes conhecidas (vs. o relatorio anterior, que teve curadoria manual):
  - Nao usa o arquivo legado "Precos antigos - Century e .pv.xlsx" (estrutura
    diferente, finishes nomeados como "OPCAO DE TOPO"); produtos cuja unica
    referencia antiga estava nesse arquivo aparecerao em "Sem Correspondente".
  - Nao reproduz a logica combinatoria especial da "Mesa Eclipse" (tampo
    inferior+superior); ela e tratada como produto comum.
  - As combinacoes podem diferir do relatorio anterior em casos ambiguos, pois
    aquele envolveu escolhas manuais (ex.: variantes de tamanho escolhidas a
    dedo, referencias cross-categoria).

Uso:
    python gerar_comparativo.py
Gera (sobrescrevendo) os dois arquivos HTML na mesma pasta.
"""
import re
import unicodedata
import warnings
from pathlib import Path

import openpyxl

warnings.filterwarnings("ignore", message="Data Validation extension is not supported")

BASE = Path(__file__).resolve().parent

NEW_PRICES_FILE = BASE / "Novos preços.xlsx"
OLD_BRAND_FILES = {
    "Century": BASE / "1 - TABELA DE PRECOS WOOD - CENTURY 28.04.xlsx",
    "Pv": BASE / "1 - TABELA DE PRECOS WOOD - PV 28.04.xlsx",
    "Private Label": BASE / "1 - TABELA DE PRECOS WOOD - PRIVATE LABEL 28.04.xlsx",
}

OUTPUT_MAIN = BASE / "Comparativo Preços Wood.html"
OUTPUT_ESP = BASE / "Comparativo Preços Wood - Mármore ESP.html"

# categoria (nome usado no relatorio) -> nome da aba nas planilhas antigas
CATEGORY_SHEETS = {
    "Mesa de Jantar": "MESA DE JANTAR",
    "Mesa de Centro": "MESA DE CENTRO",
    "Mesa de Cabeceira": "MESA DE CABECEIRA",
    "Mesa Lateral": "MESA LATERAL",
    "Aparador": "APARADOR",
    "Buffet": "BUFFET",
    "Cômoda": "CÔMODA",
}
CATEGORY_ORDER = list(CATEGORY_SHEETS)

# "MODULACAO DO PRODUTO" (planilha de precos novos) -> categoria do relatorio
MODULACAO_TO_CATEGORY = {
    "MESA JANTAR": "Mesa de Jantar",
    "MESA CENTRO": "Mesa de Centro",
    "MESA CABECEIRA": "Mesa de Cabeceira",
    "MESA LATERAL": "Mesa Lateral",
    "APARADOR": "Aparador",
    "BUFFET": "Buffet",
    "COMODA": "Cômoda",
}

# indices fixos das colunas de acabamento nas planilhas antigas (linha de cabecalho
# = "Imagem | Nome | C | L | A | (vazia) | <5 colunas de tampo> ")
COL_VIDRO_ESPELHO = 6     # Tampo: Vidro/Espelho ...
COL_LACA_VIDRO_NORMAL = 7  # Tampo: Laca/Lamina com ou sem Vidro Normal
COL_MARMORE_ESPECIAL = 8   # Tampo: Marmore (ou Porcelana) Especial
COL_MARMORE_NORMAL = 9     # Tampo: Marmore (ou Porcelana) Normal
HEADER_ROW = 6
DATA_START_ROW = 7

MARMORE_FORNECIDO_FACTOR = 0.92  # correcao: preco "sem topo" -8%

# acabamento novo -> (coluna antiga correspondente, fator de correcao)
FINISH_MAP = {
    "NENHUM": (COL_LACA_VIDRO_NORMAL, 1.0),
    "VIDRO": (COL_LACA_VIDRO_NORMAL, 1.0),
    "ESPELHO": (COL_VIDRO_ESPELHO, 1.0),
    "VIDRO FOSCO": (COL_VIDRO_ESPELHO, 1.0),
    "VIDRO EXTRACLEAR": (COL_VIDRO_ESPELHO, 1.0),
    "MARMORE POLIDO": (COL_MARMORE_NORMAL, 1.0),
    "MARMORE LEVIGADO": (COL_MARMORE_NORMAL, 1.0),
    "MARMORE ESCOVADO": (COL_MARMORE_NORMAL, 1.0),
    "MARMORE FORNECIDO": (COL_MARMORE_NORMAL, MARMORE_FORNECIDO_FACTOR),
}
ESP_FINISH_MAP = {
    "MARMORE POLIDO ESP": (COL_MARMORE_ESPECIAL, 1.0),
    "MARMORE LEVIGADO ESP": (COL_MARMORE_ESPECIAL, 1.0),
    "MARMORE ESCOVADO ESP": (COL_MARMORE_ESPECIAL, 1.0),
}

# ordem de exibicao das linhas de acabamento dentro de cada bloco de produto
FINISH_ORDER_MAIN = [
    "NENHUM", "VIDRO", "ESPELHO", "VIDRO FOSCO", "VIDRO EXTRACLEAR",
    "MARMORE POLIDO", "MARMORE LEVIGADO", "MARMORE ESCOVADO", "MARMORE FORNECIDO",
]
FINISH_ORDER_ESP = [
    "NENHUM", "VIDRO", "ESPELHO", "VIDRO FOSCO", "VIDRO EXTRACLEAR",
    "MARMORE POLIDO", "MARMORE LEVIGADO", "MARMORE ESCOVADO",
    "MARMORE POLIDO ESP", "MARMORE LEVIGADO ESP", "MARMORE ESCOVADO ESP",
    "MARMORE FORNECIDO",
]

# observacao exibida ao lado de certos acabamentos
FINISH_NOTES = {
    "VIDRO": "vidro normal = sem topo (base antiga)",
    "MARMORE FORNECIDO": "sem topo −8%",
}

NEUTRAL_THRESHOLD = 2.0  # |variacao %| < 2.0 => "neutral"

# palavras genericas ignoradas ao identificar o "nome-chave" de um produto
STOPWORDS = {
    "MESA", "DE", "DA", "DO", "E", "OU", "COM", "SEM", "ACABAMENTOS", "ACABAMENTO",
    "CUSTOMIZADOS", "CUSTOMIZADO", "JANTAR", "CENTRO", "CABECEIRA", "LATERAL",
    "BUFFET", "APARADOR", "COMODA", "RETANGULAR", "ORGANICA", "REDONDO",
    "MODELO", "TOTAL", "BAR",
}


# --------------------------------------------------------------------------- #
# utilidades
# --------------------------------------------------------------------------- #

def strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def core_tokens(name):
    """Tokens significativos (maiusculos, sem acento/pontuacao, sem stopwords)."""
    cleaned = strip_accents(name).upper()
    cleaned = re.sub(r"[^A-Z0-9 ]", " ", cleaned)
    tokens = [t for t in cleaned.split() if t and t not in STOPWORDS]
    return tokens


def fmt_brl(value):
    s = f"{value:,.2f}"
    s = s.replace(",", "§").replace(".", ",").replace("§", ".")
    return f"R$ {s}"


def fmt_brl_signed(value):
    sign = "+" if value >= 0 else "-"
    return f"{sign}{fmt_brl(abs(value))}"


def fmt_pct(value):
    sign = "+" if value >= 0 else ""
    return f"{sign}{value:.1f}%"


def variation_class(pct):
    if abs(pct) < NEUTRAL_THRESHOLD:
        return "neutral"
    return "up" if pct > 0 else "down"


def arrow(pct):
    return "↑" if pct >= 0 else "↓"


def fmt_size(c, l, a):
    a_str = f"{a:.2f}" if a is not None else ""
    c_str = f"{c:.2f}" if c is not None else ""
    l_str = f"{l:.2f}" if l is not None else ""
    return f"{c_str}x{l_str}x{a_str}m"


def short_name(display_name):
    if display_name.startswith("Mesa "):
        return display_name[len("Mesa "):].upper()
    return display_name.upper()


# --------------------------------------------------------------------------- #
# leitura: precos novos
# --------------------------------------------------------------------------- #

CONFIG_FIELDS = {
    "MODULACAO DO PRODUTO": "categoria_raw",
    "FORMATO DA MODULAÇÃO": "formato",
    "COMPRIMENTO DA MODULACAO": "C",
    "PROFUNDIDADE DA MODULACAO": "L",
    "ALTURA DA MODULACAO": "A",
}


def parse_config(text):
    info = {"categoria_raw": None, "formato": None, "C": None, "L": None, "A": None}
    if not text:
        return info
    cleaned = strip_accents(text).upper()
    for line in cleaned.splitlines():
        line = line.strip()
        if ":" not in line:
            continue
        key, _, val = line.partition(":")
        key, val = key.strip(), val.strip()
        if key == "MODULACAO DO PRODUTO":
            info["categoria_raw"] = val
        elif key == "FORMATO DA MODULACAO":
            info["formato"] = val
        elif key in ("COMPRIMENTO DA MODULACAO", "PROFUNDIDADE DA MODULACAO", "ALTURA DA MODULACAO"):
            m = re.search(r"[\d.]+", val)
            if m:
                num = float(m.group(0))
                if key.startswith("COMPRIMENTO"):
                    info["C"] = num
                elif key.startswith("PROFUNDIDADE"):
                    info["L"] = num
                else:
                    info["A"] = num
    return info


def load_new_products():
    """Retorna lista de produtos novos: cada um com nome de exibicao e variantes."""
    wb = openpyxl.load_workbook(NEW_PRICES_FILE, data_only=True)
    ws = wb["Export"]
    header = [c.value for c in ws[1]]
    finish_cols = [h for h in header[2:] if h]  # ordem original das colunas de acabamento

    groups = {}   # modelo_raw -> {"display": ..., "variants": [...]}
    order = []
    current = None
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        modelo_raw, config_text = row[0], row[1]
        if modelo_raw:
            current = modelo_raw
        if current is None:
            continue
        if current not in groups:
            display = current.replace("MODELO ", "", 1).title()
            groups[current] = {"display": display, "variants": []}
            order.append(current)

        info = parse_config(config_text)
        finishes = {}
        for name, value in zip(header[2:], row[2:]):
            if name and value is not None:
                finishes[name] = float(value)
        groups[current]["variants"].append({
            "categoria_raw": info["categoria_raw"],
            "formato": info["formato"],
            "C": info["C"], "L": info["L"], "A": info["A"],
            "finishes": finishes,
        })

    products = []
    for key in order:
        g = groups[key]
        # agrupa variantes por categoria (um mesmo modelo pode existir em > 1 categoria)
        by_cat = {}
        cat_order = []
        for v in g["variants"]:
            raw = v["categoria_raw"] or "?"
            if raw not in by_cat:
                by_cat[raw] = []
                cat_order.append(raw)
            by_cat[raw].append(v)
        for raw in cat_order:
            products.append({
                "display": g["display"],
                "categoria_raw": raw,
                "categoria": MODULACAO_TO_CATEGORY.get(raw),
                "variants": by_cat[raw],
            })
    return products, finish_cols


# --------------------------------------------------------------------------- #
# leitura: precos antigos (3 planilhas de marca)
# --------------------------------------------------------------------------- #

def load_old_products():
    """Retorna dict: categoria -> lista de produtos antigos
    {brand, name, C, L, A, cols: [v6..v10]}"""
    by_category = {cat: [] for cat in CATEGORY_SHEETS}
    for brand, path in OLD_BRAND_FILES.items():
        wb = openpyxl.load_workbook(path, data_only=True)
        for cat, sheet_name in CATEGORY_SHEETS.items():
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=DATA_START_ROW, max_row=ws.max_row, values_only=True):
                name = row[1]
                if not name or not isinstance(name, str):
                    continue
                c, l, a = row[2], row[3], row[4]
                cols = [row[6], row[7], row[8], row[9], row[10]] if len(row) > 10 else (list(row[6:10]) + [None])
                if all(v is None for v in cols):
                    continue
                by_category[cat].append({
                    "brand": brand, "name": name.strip(),
                    "C": c, "L": l, "A": a, "cols": cols,
                })
    return by_category


def numeric_dim(value, fallback):
    """Converte dimensao para numero; usa 'fallback' quando o valor nao e numerico
    (ex.: o simbolo de diametro em mesas redondas, onde C='diametro' e L=valor)."""
    if isinstance(value, (int, float)):
        return float(value)
    return fallback


def dimension_diff(new_v, old_p):
    oc = numeric_dim(old_p["C"], fallback=old_p["L"] if isinstance(old_p["L"], (int, float)) else 0.0)
    ol = old_p["L"] if isinstance(old_p["L"], (int, float)) else 0.0
    oa = old_p["A"] if isinstance(old_p["A"], (int, float)) else 0.0
    nc = new_v["C"] or 0.0
    nl = new_v["L"] or 0.0
    na = new_v["A"] or 0.0
    return abs(nc - oc) + abs(nl - ol) + abs(na - oa)


# --------------------------------------------------------------------------- #
# correspondencia (regra fixa: nome-chave + menor diferenca dimensional)
# --------------------------------------------------------------------------- #

def find_match(product, old_by_category):
    categoria = product["categoria"]
    if not categoria:
        return None
    candidates_pool = old_by_category.get(categoria, [])
    if not candidates_pool:
        return None

    new_tokens = core_tokens(product["display"])
    if not new_tokens:
        return None
    keyword = new_tokens[0]

    candidates = [p for p in candidates_pool if keyword in core_tokens(p["name"])]
    if not candidates:
        return None

    best = None
    for variant in product["variants"]:
        for old_p in candidates:
            score = dimension_diff(variant, old_p)
            if best is None or score < best[0]:
                best = (score, variant, old_p)
    if best is None:
        return None

    score, variant, old_p = best
    tied_brands = sorted({c["brand"] for c in candidates
                          if c["name"] == old_p["name"] and dimension_diff(variant, c) == score})
    brand_label = "/".join(tied_brands) if tied_brands else old_p["brand"]
    return {"variant": variant, "old": old_p, "brand_label": brand_label}


# --------------------------------------------------------------------------- #
# montagem dos dados de comparacao
# --------------------------------------------------------------------------- #

def build_rows(variant, old_p, finish_order):
    rows = []
    for finish in finish_order:
        new_price = variant["finishes"].get(finish)
        if new_price is None:
            continue
        mapping = FINISH_MAP.get(finish) or ESP_FINISH_MAP.get(finish)
        if mapping is None:
            continue
        col_idx, factor = mapping
        old_raw = old_p["cols"][col_idx - COL_VIDRO_ESPELHO]
        if not isinstance(old_raw, (int, float)):
            continue
        old_price = old_raw * factor
        delta = new_price - old_price
        pct = (delta / old_price) * 100 if old_price else 0.0
        rows.append({
            "finish": finish,
            "note": FINISH_NOTES.get(finish),
            "old": old_price, "new": new_price,
            "delta": delta, "pct": pct,
        })
    return rows


def build_report_data(finish_order):
    new_products, _ = load_new_products()
    old_by_category = load_old_products()

    matched_by_cat = {cat: [] for cat in CATEGORY_ORDER}
    unmatched = []

    for product in new_products:
        match = find_match(product, old_by_category)
        if match is None:
            unmatched.append(product)
            continue
        rows = build_rows(match["variant"], match["old"], finish_order)
        if not rows:
            unmatched.append(product)
            continue
        matched_by_cat[product["categoria"]].append({
            "display": product["display"],
            "brand_label": match["brand_label"],
            "size": fmt_size(match["variant"]["C"], match["variant"]["L"], match["variant"]["A"]),
            "ref": match["old"]["name"],
            "rows": rows,
        })

    return matched_by_cat, unmatched


# --------------------------------------------------------------------------- #
# geracao do HTML
# --------------------------------------------------------------------------- #

PAGE_CSS = """
:root{--dg:#484c40;--mg:#84867b;--kh:#878264;--sd:#c5c0b3;--ow:#dcd8d3;--cr:#f2f0ec;--bg:#ede9e4;--ch:#211f1e;--wm:#494038;--bl:#909d9c;--up:#2e6b35;--upbg:#ebf4ec;--upbd:#b8dbbf;--dn:#8b2020;--dnbg:#fdf0f0;--dnbd:#e8b4b4;--nt:#5a5849;--ntbg:#f5f3ef;--ntbd:#ccc9c0;--ep:#4a3a6b;--epbg:#f0ecfa;--epbd:#c4b4e8;--nm:#5a4520;--nmbg:#faf4eb;--nmbd:#d4b87a}
*{margin:0;padding:0;box-sizing:border-box}html{scroll-behavior:smooth}
body{font-family:'Albert Sans',sans-serif;background:var(--cr);color:var(--ch);font-size:13px;line-height:1.6}
.hdr{background:var(--ch);color:#fff;padding:52px 64px 44px;position:relative;overflow:hidden}
.hdr::after{content:'';position:absolute;top:-120px;right:-120px;width:400px;height:400px;border-radius:50%;background:rgba(255,255,255,.03);pointer-events:none}
.hdr-eye{font-size:9px;font-weight:600;letter-spacing:.3em;text-transform:uppercase;color:var(--sd);margin-bottom:14px;opacity:.8}
.hdr-title{font-family:'Geologica',sans-serif;font-size:42px;font-weight:200;letter-spacing:-.02em;line-height:1.05;margin-bottom:6px}
.hdr-title b{font-weight:700}.hdr-sub{color:var(--sd);font-size:13px;font-weight:300;opacity:.8;margin-bottom:32px}
.hdr-pills{display:flex;flex-wrap:wrap;gap:28px;border-top:1px solid rgba(255,255,255,.1);padding-top:24px}
.pill{display:flex;flex-direction:column;gap:3px}.pill-l{font-size:8px;letter-spacing:.22em;text-transform:uppercase;color:rgba(255,255,255,.4)}.pill-v{font-size:13px;font-weight:500;color:rgba(255,255,255,.85)}
.hdr-brand{position:absolute;right:64px;bottom:44px;font-family:'Geologica',sans-serif;font-size:13px;font-weight:200;letter-spacing:.4em;text-transform:uppercase;color:rgba(255,255,255,.18)}.hdr-brand b{font-weight:700;color:rgba(255,255,255,.32)}
.nav{position:sticky;top:0;z-index:200;background:var(--dg);display:flex;overflow-x:auto;scrollbar-width:none;border-bottom:1px solid rgba(0,0,0,.2)}.nav::-webkit-scrollbar{display:none}
.nav-home{padding:13px 18px 13px 24px;font-family:'Geologica',sans-serif;font-size:10px;font-weight:700;letter-spacing:.25em;text-transform:uppercase;color:rgba(255,255,255,.55);text-decoration:none;white-space:nowrap;border-right:1px solid rgba(255,255,255,.1);margin-right:8px}
.nav-a{padding:13px 16px;font-size:10px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:rgba(255,255,255,.45);text-decoration:none;white-space:nowrap;border-bottom:2px solid transparent;transition:all .2s}.nav-a:hover{color:#fff;border-bottom-color:var(--sd)}
.nav-a.nm-nav{color:rgba(212,184,122,.6)}.nav-a.nm-nav:hover{color:#d4b87a;border-bottom-color:#d4b87a}
.main{max-width:1180px;margin:0 auto;padding:52px 64px 96px}
.eyebrow{font-family:'Geologica',sans-serif;font-size:9px;font-weight:600;letter-spacing:.28em;text-transform:uppercase;color:var(--kh);margin-bottom:18px}
.note{background:#fff;border:1px solid var(--ow);border-left:3px solid var(--kh);padding:14px 18px;margin-bottom:36px;font-size:12px;color:var(--wm);line-height:1.8}.note b{font-weight:600;color:var(--dg)}
.cgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(172px,1fr));gap:10px;margin-bottom:60px}
.ccard{background:#fff;border:1px solid var(--ow);padding:20px 18px 16px;cursor:pointer;transition:border-color .2s,box-shadow .2s}.ccard:hover{border-color:var(--sd);box-shadow:0 2px 12px rgba(0,0,0,.06)}
.ccard-nm{font-family:'Geologica',sans-serif;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:var(--dg);margin-bottom:1px}.ccard-sub{font-size:10px;color:var(--mg);margin-bottom:14px}
.ccard-avg{font-family:'Geologica',sans-serif;font-size:28px;font-weight:200;line-height:1;margin-bottom:2px}.ccard-avg.up{color:var(--up)}.ccard-avg.down{color:var(--dn)}.ccard-avg.neutral{color:var(--nt)}
.ccard-lbl{font-size:9px;letter-spacing:.12em;text-transform:uppercase;color:var(--mg);margin-bottom:12px}
.ccard-rng{display:flex;flex-wrap:wrap;align-items:flex-start;gap:4px;font-size:10px}.ccard-rng span{display:flex;flex-direction:column;gap:1px}.ccard-rng em{font-style:normal;font-size:9px;color:var(--mg)}.ccard-rng .up{color:var(--up)}.ccard-rng .down{color:var(--dn)}.ccard-rng .neutral{color:var(--nt)}.sep{color:var(--ow);padding:0 3px;align-self:center}
.csec{margin-bottom:60px}.csec-hdr{display:flex;align-items:center;gap:14px;padding-bottom:14px;margin-bottom:20px;border-bottom:2px solid var(--ow)}.csec-hdr h2{font-family:'Geologica',sans-serif;font-size:22px;font-weight:300;color:var(--ch);letter-spacing:-.01em}
.cbadge{font-size:11px;font-weight:600;padding:3px 11px;border-radius:20px}.cbadge.up{background:var(--upbg);color:var(--up);border:1px solid var(--upbd)}.cbadge.down{background:var(--dnbg);color:var(--dn);border:1px solid var(--dnbd)}.cbadge.neutral{background:var(--ntbg);color:var(--nt);border:1px solid var(--ntbd)}
.nm-intro{font-size:12px;color:var(--nm);background:var(--nmbg);border:1px solid var(--nmbd);padding:10px 16px;margin-bottom:16px;border-radius:2px}
.pblock{background:#fff;border:1px solid var(--ow);margin-bottom:12px;overflow:hidden}
.ph{padding:12px 18px 10px;background:var(--bg);border-bottom:1px solid var(--ow)}
.pname{font-family:'Geologica',sans-serif;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:var(--ch);margin-bottom:5px}
.pmeta{display:flex;flex-wrap:wrap;gap:5px;align-items:center}.tag{font-size:9px;font-weight:600;padding:2px 8px;letter-spacing:.05em}.tag.brand{background:var(--ch);color:rgba(255,255,255,.6)}.tag.size{background:var(--dg);color:rgba(255,255,255,.7)}.tag.ref{font-weight:400;font-style:italic;background:transparent;border:1px solid var(--ow);color:var(--mg)}
.ft{width:100%;border-collapse:collapse}.ft thead th{background:var(--ch);color:rgba(255,255,255,.5);font-size:9px;font-weight:600;letter-spacing:.18em;text-transform:uppercase;padding:8px 14px;text-align:left}.th-pr,.th-vr,.th-ab{text-align:right!important}
.ft tbody tr{border-bottom:1px solid var(--ow)}.ft tbody tr:last-child{border:none}.ft tbody tr:hover{background:var(--cr)}.ft td{padding:9px 14px;vertical-align:middle}
.fn{font-size:12px;color:var(--ch);min-width:160px}.fn-note{font-size:9px;color:var(--bl);font-style:italic;margin-left:6px}
.pr{text-align:right;font-variant-numeric:tabular-nums;font-size:12px}.pr.old{color:var(--mg)}.pr.new{color:var(--ch);font-weight:500}
.vr{text-align:right;font-size:12px;font-weight:700;min-width:80px}.ab{text-align:right;font-size:11px;min-width:100px}
.up{color:var(--up)}.down{color:var(--dn)}.neutral{color:var(--nt)}
.pblock.no-match{border-color:var(--nmbd);border-left:3px solid var(--nmbd)}.pblock.no-match .ph{background:var(--nmbg)}.nm-badge{font-size:9px;font-weight:500;background:var(--nmbg);border:1px solid var(--nmbd);color:var(--nm);padding:2px 8px;border-radius:3px;letter-spacing:.04em;vertical-align:middle;margin-left:8px}
.ftr{background:var(--ch);color:rgba(255,255,255,.3);padding:24px 64px;display:flex;justify-content:space-between;align-items:center;font-size:11px}.ftr-brand{font-family:'Geologica',sans-serif;letter-spacing:.3em;text-transform:uppercase;font-weight:200;color:rgba(255,255,255,.2)}.ftr-brand b{font-weight:700}
"""

NAV_LINKS = "".join(
    f'<a class="nav-a" href="#sec-{cat.lower().replace(" ", "-").replace("ê", "e").replace("ô", "o")}">{cat}</a>'
    for cat in CATEGORY_ORDER
)


def render_summary_cards(matched_by_cat):
    cards = []
    for cat in CATEGORY_ORDER:
        products = matched_by_cat[cat]
        if not products:
            continue
        all_rows = [(r, p["display"]) for p in products for r in p["rows"]]
        pcts = [r["pct"] for r, _ in all_rows]
        avg = sum(pcts) / len(pcts)
        worst = min(all_rows, key=lambda x: x[0]["pct"])
        best = max(all_rows, key=lambda x: x[0]["pct"])
        unit = "produto" if len(products) == 1 else "produtos"
        cards.append(f"""<div class="ccard" onclick="go('{cat}')">
  <div class="ccard-nm">{cat}</div><div class="ccard-sub">{len(products)} {unit}</div>
  <div class="ccard-avg {variation_class(avg)}">{fmt_pct(avg)}</div><div class="ccard-lbl">variação média</div>
  <div class="ccard-rng">
    <span class="{variation_class(worst[0]['pct'])}">{fmt_pct(worst[0]['pct'])} <em>{short_name(worst[1])}</em></span>
    <span class="sep">·</span>
    <span class="{variation_class(best[0]['pct'])}">{fmt_pct(best[0]['pct'])} <em>{short_name(best[1])}</em></span>
  </div>
</div>""")
    return "".join(cards)


def render_product_block(product):
    rows_html = []
    for r in product["rows"]:
        note = f'<span class="fn-note">{r["note"]}</span>' if r["note"] else ""
        rows_html.append(f"""<tr>
  <td class="fn">{r['finish']}{note}</td>
  <td class="pr old">{fmt_brl(r['old'])}</td><td class="pr new">{fmt_brl(r['new'])}</td>
  <td class="vr {variation_class(r['pct'])}">{arrow(r['pct'])} {fmt_pct(r['pct'])}</td>
  <td class="ab {'up' if r['delta'] >= 0 else 'down'}">{fmt_brl_signed(r['delta'])}</td></tr>""")
    return f"""<div class="pblock">
  <div class="ph"><div class="pname">{product['display']}</div>
  <div class="pmeta"><span class="tag brand">{product['brand_label']}</span><span class="tag size">{product['size']}</span><span class="tag ref">{product['ref']}</span></div></div>
  <table class="ft"><thead><tr>
    <th class="th-fn">Acabamento</th>
    <th class="th-pr">Preço Anterior</th><th class="th-pr">Preço Novo</th>
    <th class="th-vr">Var %</th><th class="th-ab">Var R$</th>
  </tr></thead><tbody>{''.join(rows_html)}</tbody></table>
</div>"""


def render_category_section(cat, products):
    if not products:
        return ""
    all_pcts = [r["pct"] for p in products for r in p["rows"]]
    avg = sum(all_pcts) / len(all_pcts)
    slug = cat.lower().replace(" ", "-").replace("ê", "e").replace("ô", "o")
    blocks = "".join(render_product_block(p) for p in products)
    return f"""<section class="csec" id="sec-{slug}"><div class="csec-hdr"><h2>{cat}</h2><span class="cbadge {variation_class(avg)}">{fmt_pct(avg)} média</span></div>{blocks}</section>"""


def render_unmatched(unmatched, finish_cols):
    if not unmatched:
        return ""
    blocks = []
    for product in unmatched:
        rows = []
        for variant in product["variants"]:
            for finish in finish_cols:
                price = variant["finishes"].get(finish)
                if price is None:
                    continue
                rows.append(f'<tr><td class="fn">{finish}</td><td class="pr new" colspan="4">{fmt_brl(price)}</td></tr>')
        if not rows:
            continue
        first = product["variants"][0]
        size = fmt_size(first["C"], first["L"], first["A"])
        blocks.append(f"""<div class="pblock no-match">
  <div class="ph"><div class="pname">{product['display']} <span class="nm-badge">Sem correspondente — apenas listagem</span></div>
  <div class="pmeta"><span class="tag brand">—</span><span class="tag size">{size}</span></div></div>
  <table class="ft"><thead><tr><th class="th-fn">Acabamento</th><th class="th-pr" colspan="4">Preço Novo</th></tr></thead><tbody>{''.join(rows)}</tbody></table>
</div>""")
    return f"""<section class="csec" id="sec-sem-correspondente">
  <div class="csec-hdr"><h2>Sem Correspondente</h2><span class="cbadge neutral">apenas listagem</span></div>
  <div class="nm-intro">Produtos abaixo não possuem equivalente nas tabelas de preços antigas. Os valores novos são listados para referência.</div>
  {''.join(blocks)}</section>"""


def render_page(matched_by_cat, unmatched, finish_cols, *, esp_variant):
    total_matched = sum(len(v) for v in matched_by_cat.values())
    total_unmatched = len(unmatched)

    if esp_variant:
        title_extra = " - Mármore ESP"
        esp_note = ""
        nav_unmatched_class = ""
    else:
        title_extra = ""
        esp_note = '\n    <b>Mármores Especiais (ESP)</b> → <b>excluídos das comparações</b> nesta versão.&nbsp;'
        nav_unmatched_class = " nm-nav"

    note = f"""<div class="note">
    <b>Correções aplicadas:</b>&nbsp;
    <b>Mármore Fornecido</b> → preço sem topo −8%.&nbsp;
    <b>Vidro Normal</b> → mesmo preço do sem topo (coluna "Tampo: Laca ou Lamina com ou sem Vidro Normal").&nbsp;{esp_note}
    <b>Sem correspondente</b> → agrupados ao final, apenas listagem de preços novos.
  </div>"""

    sections = "".join(render_category_section(cat, matched_by_cat[cat]) for cat in CATEGORY_ORDER)

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Análise Comparativa de Preços{title_extra} — Wood | SoHome</title>
<link href="https://fonts.googleapis.com/css2?family=Geologica:wght@200;300;400;600;700&family=Albert+Sans:ital,wght@0,300;0,400;0,500;0,600;1,300;1,400&display=swap" rel="stylesheet">
<style>{PAGE_CSS}</style>
</head>
<body>
<header class="hdr">
  <div class="hdr-eye">Grupo SoHome · Linha Wood</div>
  <h1 class="hdr-title">Análise Comparativa<br><b>de Preços</b></h1>
  <p class="hdr-sub">Reprecificação vs. Tabelas 28/04 — por categoria e acabamento</p>
  <div class="hdr-pills">
    <div class="pill"><span class="pill-l">Com comparativo</span><span class="pill-v">{total_matched} produtos</span></div>
    <div class="pill"><span class="pill-l">Sem correspondente</span><span class="pill-v">{total_unmatched} produtos</span></div>
    <div class="pill"><span class="pill-l">Fontes antigas</span><span class="pill-v">Century · PV · Private Label 28/04</span></div>
  </div>
  <div class="hdr-brand">SO<b>HOME</b></div>
</header>
<nav class="nav">
  <a class="nav-home" href="#top">SO<b>HOME</b></a>
  <a class="nav-a" href="#resumo">Resumo</a>{NAV_LINKS}<a class="nav-a{nav_unmatched_class}" href="#sec-sem-correspondente">Sem Correspondente</a>
</nav>
<main class="main" id="top">
<section id="resumo" style="margin-bottom:60px">
  <p class="eyebrow">Resumo por Categoria</p>
  {note}
  <div class="cgrid">{render_summary_cards(matched_by_cat)}</div>
</section>
<section><p class="eyebrow">Detalhamento por Produto</p>{sections}{render_unmatched(unmatched, finish_cols)}</section>
</main>
<footer class="ftr">
  <span class="ftr-brand">SO<b>HOME</b> · Wood · Análise de Preços</span>
  <span>Comparativo: Reprecificação vs. Tabelas 28/04</span>
</footer>
<script>function go(cat){{const s=cat.toLowerCase().replace(/ /g,'-').replace(/ê/g,'e').replace(/ô/g,'o');const el=document.getElementById('sec-'+s);if(el)el.scrollIntoView({{behavior:'smooth',block:'start'}});}}</script>
</body></html>
"""


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #

def main():
    _, finish_cols = load_new_products()

    print("Lendo planilhas e calculando comparações (sem ESP)...")
    matched, unmatched = build_report_data(FINISH_ORDER_MAIN)
    OUTPUT_MAIN.write_text(render_page(matched, unmatched, finish_cols, esp_variant=False), encoding="utf-8")
    print(f"  -> {OUTPUT_MAIN.name}  "
          f"({sum(len(v) for v in matched.values())} com comparativo, {len(unmatched)} sem correspondente)")

    print("Calculando versão com Mármore ESP...")
    matched_esp, unmatched_esp = build_report_data(FINISH_ORDER_ESP)
    OUTPUT_ESP.write_text(render_page(matched_esp, unmatched_esp, finish_cols, esp_variant=True), encoding="utf-8")
    print(f"  -> {OUTPUT_ESP.name}  "
          f"({sum(len(v) for v in matched_esp.values())} com comparativo, {len(unmatched_esp)} sem correspondente)")


if __name__ == "__main__":
    main()
