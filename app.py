import re
import os
import io
import unicodedata
from math import ceil
from datetime import datetime

from pypdf import PdfReader
import streamlit as st
import xlsxwriter

# =========================
# Config
# =========================
st.set_page_config(page_title="PDF → Excel (Lite)", page_icon="🪶", layout="wide")
st.title("🪶 PDF → Excel (Lite)")
st.caption("Parser robusto — nomes limpos, números do Lince (3.491.40), busca, paginação, seleção persistente. (pypdf + xlsxwriter)")

# -------------------------
# Utilidades
# -------------------------
def br_to_float(txt: str):
    if txt is None:
        return None
    t = txt.strip()
    if not t:
        return None
    if "," in t:
        try:
            return float(t.replace(".", "").replace(",", "."))
        except Exception:
            return None
    if t.count(".") >= 2:
        parts = t.split(".")
        intpart = "".join(parts[:-1])
        dec = parts[-1]
        try:
            return float(intpart + "." + dec)
        except Exception:
            return None
    try:
        return float(t)
    except Exception:
        return None

def is_num_token(tok: str) -> bool:
    return bool(re.fullmatch(r"[0-9][0-9\.\,]*", tok or ""))

def dec_places(tok: str) -> int:
    if not tok:
        return 0
    s = tok.replace(".", ",")
    if "," in s:
        return len(s.split(",")[-1])
    return 0

def extract_text_with_pypdf(file) -> str:
    reader = PdfReader(file)
    texts = []
    for page in reader.pages:
        try:
            texts.append(page.extract_text() or "")
        except Exception:
            texts.append("")
    return "\n".join(texts)

SETORES_CANON = [
    "Frios", "Padaria", "Confeitaria Fina", "Confeitaria Trad",
    "Restaurante", "Salgados", "Lanchonete"
]

def _norm(s: str) -> str:
    s = unicodedata.normalize("NFD", s or "").encode("ascii", "ignore").decode("ascii")
    return s.upper()

def guess_setor(text: str, filename: str) -> str:
    hay = _norm((text or "") + " " + (filename or ""))
    if any(k in hay for k in ["FRIO", "FIOS"]):        return "Frios"
    if "PADARIA" in hay:                               return "Padaria"
    if "CONFEITARIA FINA" in hay or "FINA" in hay:     return "Confeitaria Fina"
    if "CONFEITARIA TRAD" in hay or "TRAD" in hay:     return "Confeitaria Trad"
    if "RESTAURANTE" in hay:                           return "Restaurante"
    if "SALGADOS" in hay:                              return "Salgados"
    if "LANCHONETE" in hay:                            return "Lanchonete"
    return "Frios"

def glue_wrapped_lines(lines):
    glued = []
    i = 0
    while i < len(lines):
        cur = lines[i]
        nxt = lines[i+1] if i + 1 < len(lines) else ""
        cur_toks = cur.split()
        nxt_toks = nxt.split()

        j = len(cur_toks)
        while j > 0 and is_num_token(cur_toks[j-1]):
            j -= 1
        cur_tail_len = len(cur_toks) - j
        nxt_num_ratio = (sum(1 for t in nxt_toks if is_num_token(t)) / max(1, len(nxt_toks))) if nxt_toks else 0.0

        if cur_tail_len < 2 and nxt_num_ratio >= 0.5:
            glued.append((cur + " " + nxt).strip())
            i += 2
        else:
            glued.append(cur)
            i += 1
    return glued

def clean_tokens(tokens):
    out = []
    removed_leading_code = False
    for idx, t in enumerate(tokens):
        if re.fullmatch(r"\d{12,}", t):
            continue
        if not removed_leading_code and idx == 0 and re.fullmatch(r"\d{3,6}", t):
            removed_leading_code = True
            continue
        out.append(t)
    return out

def parse_lince_lines_to_list(text: str):
    lines = [re.sub(r"\s{2,}", " ", (ln or "")).strip() for ln in text.splitlines()]
    lixo = ("Curva ABC","Período","CST","ECF","Situação Tributária","Classif.","Codigo","CÓDIGO",
            "Barras","Total do Departamento","Total Geral","www.grupotecnoweb.com.br")
    lines = [ln for ln in lines if ln and not any(k in ln for k in lixo)]

    cleaned = []
    for ln in lines:
        ln = re.sub(r"\b\d{8,13}\b\s*$", "", ln).strip()
        ln = re.sub(r"\b\d{4,8}\b\s*$", "", ln).strip()
        cleaned.append(ln)
    cleaned = glue_wrapped_lines(cleaned)

    items_raw = []
    for ln in cleaned:
        toks = ln.split()
        if not toks:
            continue
        toks = clean_tokens(toks)
        if not toks:
            continue

        idx = len(toks)
        while idx > 0 and is_num_token(toks[idx-1]):
            idx -= 1
        head = toks[:idx]
        tail = toks[idx:]
        if len(tail) < 2 or not head:
            continue

        i_qtd = None
        for j in range(len(tail)-1, -1, -1):
            if dec_places(tail[j]) == 3 and br_to_float(tail[j]) is not None:
                i_qtd = j
                break
        i_val = None
        if i_qtd is not None:
            for j in range(i_qtd+1, len(tail)):
                if dec_places(tail[j]) == 2 and br_to_float(tail[j]) is not None:
                    i_val = j
                    break

        if i_qtd is None or i_val is None:
            qtd = br_to_float(tail[-2]); valor = br_to_float(tail[-1])
        else:
            qtd = br_to_float(tail[i_qtd]); valor = br_to_float(tail[i_val])

        if qtd is None or valor is None or qtd < 0 or valor < 0:
            continue

        head_clean = [t for t in head if not is_num_token(t)]
        nome = re.sub(r"\s{2,}", " ", " ".join(head_clean)).strip()
        if not re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", nome):
            continue

        items_raw.append({"nome": nome, "quantidade": float(qtd), "valor": float(valor)})

    agg = {}
    for it in items_raw:
        k = it["nome"]
        if k not in agg:
            agg[k] = {"nome": k, "quantidade": 0.0, "valor": 0.0}
        agg[k]["quantidade"] += it["quantidade"]
        agg[k]["valor"] += it["valor"]

    return sorted(agg.values(), key=lambda x: x["valor"], reverse=True)

# -------------------------
# Inputs
# -------------------------
uploaded = st.file_uploader("Envie o PDF (Curva ABC do Lince)", type=["pdf"])
default_mes = datetime.today().strftime("%m/%Y")
mes = st.text_input("Mês (ex.: 08/2025)", value=default_mes)
semana = st.text_input("Semana (ex.: 1ª semana de ago/2025)", value="")

# -------------------------
# UI + Geração
# -------------------------
if uploaded:
    all_text = extract_text_with_pypdf(uploaded)
    setor_guess = guess_setor(all_text, uploaded.name)
    try:
        idx = SETORES_CANON.index(setor_guess)
    except ValueError:
        idx = 0
    setor = st.selectbox("Setor", SETORES_CANON, index=idx)

    rows_all = parse_lince_lines_to_list(all_text)
    if not rows_all:
        st.error("Não consegui identificar linhas de produto neste PDF.")
        st.code(all_text[:2000]); st.stop()

    q = st.text_input("🔎 Buscar produto (contém):", value="").strip().upper()
    rows = [r for r in rows_all if q in r["nome"].upper()] if q else rows_all[:]

    order = st.selectbox("Ordenar por", ["valor (desc)", "quantidade (desc)", "nome (A→Z)"], index=0)
    if order.startswith("valor"):
        rows.sort(key=lambda x: x["valor"], reverse=True)
    elif order.startswith("quantidade"):
        rows.sort(key=lambda x: x["quantidade"], reverse=True)
    else:
        rows.sort(key=lambda x: x["nome"])

    page_size = st.selectbox("Itens por página", [20, 50, 100], index=0)
    total = len(rows); pages = max(1, ceil(total / page_size))
    page = st.number_input("Página", min_value=1, max_value=pages, value=1, step=1)
    start = (page - 1) * page_size; end = start + page_size
    page_rows = rows[start:end]

    if "selecao" not in st.session_state:
        st.session_state.selecao = {}
    for r in rows_all:
        st.session_state.selecao.setdefault(r["nome"], True)

    st.markdown("---")
    for r in page_rows:
        nome = r["nome"]; qtd = round(float(r["quantidade"]), 3); val = round(float(r["valor"]), 2)
        cols = st.columns([0.6, 4.0, 1.4, 1.4])
        st.session_state.selecao[nome] = cols[0].checkbox("", value=st.session_state.selecao.get(nome, True), key=f"chk_{nome}")
        cols[1].text(nome)
        cols[2].text(f"{qtd:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
        cols[3].text(f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    if st.button("Gerar Excel (.xlsx)"):
        selecionados = sorted(
            [r for r in rows_all if st.session_state.selecao.get(r['nome'], False)],
            key=lambda x: x["valor"], reverse=True
        )
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = workbook.add_worksheet("Produtos")
        headers = ["nome do produto", "setor", "mês", "semana", "quantidade", "valor"]
        for col, h in enumerate(headers):
            ws.write(0, col, h)
        for i, r in enumerate(selecionados, start=1):
            ws.write(i, 0, r["nome"]); ws.write(i, 1, setor); ws.write(i, 2, mes); ws.write(i, 3, semana)
            ws.write_number(i, 4, round(float(r["quantidade"]), 3))
            ws.write_number(i, 5, round(float(r["valor"]), 2))
        workbook.close()
        st.download_button("⬇️ Baixar Excel", data=output.getvalue(), file_name=f"produtos_{mes.replace('/', '-')}.xlsx")
else:
    st.info("Envie um PDF para começar.")

