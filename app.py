# app.py final com busca múltipla
import re
import os
import io
from math import ceil
from datetime import datetime

from pypdf import PdfReader
import streamlit as st
import xlsxwriter

st.set_page_config(page_title="PDF → Excel (Lite)", page_icon="🪶", layout="wide")
st.title("🪶 PDF → Excel (Lite)")

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
    return bool(re.fullmatch(r"[0-9][0-9\.,]*", tok or ""))

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

def guess_setor(text: str, filename: str) -> str:
    base = os.path.basename(filename or "")
    base_up = base.upper()
    for chave in ["FRIOS","ACOUGUE","AÇOUGUE","PADARIA","HORTIFRUTI","BEBIDAS","MERCEARIA","LANCHONETE"]:
        if chave in base_up:
            return chave
    return "N/D"

def parse_lince_lines_to_list(text: str):
    lines = [re.sub(r"\s{2,}", " ", (ln or "")).strip() for ln in text.splitlines()]
    lixo = ("Curva ABC","Período","CST","ECF","Situação Tributária","Classif.","Codigo","CÓDIGO",
            "Barras","Total do Departamento","Total Geral","www.grupotecnoweb.com.br")
    lines = [ln for ln in lines if ln and not any(k in ln for k in lixo)]
    items_raw = []
    for ln in lines:
        toks = ln.split()
        if not toks:
            continue
        head = []
        tail = []
        for t in toks:
            if is_num_token(t):
                tail.append(t)
            else:
                head.append(t)
        if len(tail) < 2:
            continue
        qtd = br_to_float(tail[-2]); valor = br_to_float(tail[-1])
        if qtd is None or valor is None:
            continue
        nome = " ".join(head)
        items_raw.append({"nome": nome, "quantidade": float(qtd), "valor": float(valor)})
    return sorted(items_raw, key=lambda x: x["valor"], reverse=True)

def tokenize_multi_query(raw: str):
    raw = (raw or "").strip()
    if not raw:
        return [], [], []
    exact = re.findall(r'"([^"]+)"', raw)
    tmp = re.sub(r'"[^"]+"', " ", raw)
    parts = []
    for chunk in re.split(r'[,;\n]+', tmp):
        chunk = chunk.strip()
        if chunk:
            parts.append(chunk)
    includes, excludes = [], []
    for p in parts:
        if p.startswith("-") and len(p) > 1:
            excludes.append(p[1:].strip())
        else:
            includes.append(p)
    exact   = [t.upper() for t in exact]
    includes = [t.upper() for t in includes]
    excludes = [t.upper() for t in excludes]
    return includes, excludes, exact

def name_matches(name: str, includes, excludes, exact, require_all: bool) -> bool:
    N = (name or "").upper()
    for ex in exact:
        if ex not in N:
            return False
    if includes:
        if require_all:
            for inc in includes:
                if inc not in N:
                    return False
        else:
            if not any(inc in N for inc in includes):
                return False
    for exc in excludes:
        if exc in N:
            return False
    return True

uploaded = st.file_uploader("Envie o PDF (Curva ABC do Lince)", type=["pdf"])
default_mes = datetime.today().strftime("%m/%Y")
mes = st.text_input("Mês (ex.: 08/2025)", value=default_mes)
semana = st.text_input("Semana", value="")

if uploaded:
    all_text = extract_text_with_pypdf(uploaded)
    setor_guess = guess_setor(all_text, uploaded.name)
    setor = st.text_input("Setor", value=setor_guess)
    rows_all = parse_lince_lines_to_list(all_text)
    if not rows_all:
        st.error("Não consegui identificar linhas de produto neste PDF.")
        st.stop()

    st.markdown("### 🔎 Buscar produtos (múltiplos)")
    q_raw = st.text_area("Digite termos separados por vírgula ou linha", value="", height=90)
    mode = st.radio("Modo", ["Qualquer termo (OR)", "Todos os termos (AND)"], index=0, horizontal=True)
    preselect = st.checkbox("Pré-selecionar resultados da busca", value=False)

    inc, exc, exa = tokenize_multi_query(q_raw)
    require_all = (mode == "Todos os termos (AND)")
    if any([inc, exc, exa]):
        rows = [r for r in rows_all if name_matches(r["nome"], inc, exc, exa, require_all)]
    else:
        rows = rows_all[:]
    if preselect:
        if "selecao" not in st.session_state:
            st.session_state.selecao = {}
        for r in rows:
            st.session_state.selecao[r["nome"]] = True

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

    for r in page_rows:
        nome = r["nome"]
        qtd = round(float(r["quantidade"]), 3)
        val = round(float(r["valor"]), 2)
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
        if not selecionados:
            st.warning("Selecione ao menos um produto."); st.stop()
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
        st.download_button(
            label="⬇️ Baixar Excel",
            data=output.getvalue(),
            file_name=f"produtos_{mes.replace('/', '-')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Envie um PDF para começar.")
