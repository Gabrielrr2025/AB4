import io
import re
import os
from math import ceil
from datetime import datetime

from pypdf import PdfReader
import streamlit as st
import xlsxwriter

# =========================
# Config da página
# =========================
st.set_page_config(page_title="PDF → Excel (Lite)", page_icon="🪶", layout="wide")
st.title("🪶 PDF → Excel (Lite)")
st.caption("Sem pandas/pdfplumber — pypdf + xlsxwriter. Busca, paginação, checkboxes e Top N por valor.")

NUM_TOKEN = r"[0-9\.\,]+"

# -------------------------
# Utilidades
# -------------------------
def br_to_float(txt: str):
    """Converte '1.234,56' → 1234.56; e '1,234.56' → 1234.56."""
    if txt is None:
        return None
    t = txt.strip()
    if not t:
        return None
    # tenta BR
    if "," in t:
        t1 = t.replace(".", "").replace(",", ".")
        try:
            return float(t1)
        except Exception:
            pass
    # tenta EN
    t2 = t.replace(",", "")
    try:
        return float(t2)
    except Exception:
        return None

def guess_setor(text: str, filename: str) -> str:
    """Tenta achar setor no texto ou deduzir pelo nome do arquivo."""
    m = re.search(r"Departamento:\s*([\s\S]{0,60})", text, flags=re.IGNORECASE)
    if m:
        tail = text[m.end():].splitlines()
        for ln in tail[:5]:
            t = (ln or "").strip()
            if 2 <= len(t) <= 25 and t.upper() == t:
                return t
    base = os.path.basename(filename or "")
    base_up = base.upper()
    for chave in ["FRIOS", "ACOUGUE", "AÇOUGUE", "PADARIA", "HORTIFRUTI", "BEBIDAS", "MERCEARIA", "LANCHONETE"]:
        if chave in base_up:
            start = base_up.find(chave)
            end = min(len(base_up), start + len(chave) + 2)
            return re.sub(r"[^A-Z0-9]", "", base_up[start:end])
    return "N/D"

def extract_text_with_pypdf(file) -> str:
    """Extrai texto de todas as páginas (tolerante a erros)."""
    reader = PdfReader(file)
    texts = []
    for page in reader.pages:
        try:
            texts.append(page.extract_text() or "")
        except Exception:
            texts.append("")
    return "\n".join(texts)

def is_num_token(tok: str) -> bool:
    return re.fullmatch(r"[0-9][0-9\.\,]*", tok or "") is not None

def glue_wrapped_lines(lines):
    """
    Une linhas que foram quebradas no PDF: se uma linha NÃO tem tail numérico suficiente
    (>=2 números no final) e a próxima linha é majoritariamente numérica, concatena.
    """
    glued = []
    i = 0
    while i < len(lines):
        cur = lines[i]
        nxt = lines[i+1] if i + 1 < len(lines) else ""
        cur_toks = cur.split()
        nxt_toks = nxt.split()

        # identifica tail numérico em cur (contíguo no fim)
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

def parse_lince_lines_to_list(text: str):
    """
    Extrai itens do relatório 'Curva ABC' (Lince).
    Estratégia:
      - Normaliza espaços e remove cabeçalhos/rodapés.
      - Cola linhas quebradas (glue_wrapped_lines).
      - Para cada linha, separa tokens e identifica o TAIL numérico contíguo no fim.
      - Usa os DOIS últimos números do TAIL como (quantidade, valor) (ordem ajustada abaixo),
        e, se existir, o número anterior como preço (não usamos).
      - Nome = tudo ANTES do TAIL (garante que números “soltos” não contaminem o nome).
      - Agrega por nome e ordena por 'valor' desc.
    Retorna: lista de dicts {"nome", "quantidade", "valor"}.
    """
    # 1) normaliza
    lines = [re.sub(r"\s{2,}", " ", (ln or "")).strip() for ln in text.splitlines()]
    lixo = (
        "Curva ABC", "Período", "CST", "ECF", "Situação Tributária",
        "Classif.", "Codigo", "CÓDIGO", "Barras", "Total do Departamento",
        "Total Geral", "www.grupotecnoweb.com.br"
    )
    lines = [ln for ln in lines if ln and not any(k in ln for k in lixo)]

    # 2) remove EAN/código no final e cola linhas quebradas
    cleaned = []
    for ln in lines:
        ln = re.sub(r"\b\d{8,13}\b\s*$", "", ln).strip()  # EAN
        ln = re.sub(r"\b\d{4,8}\b\s*$", "", ln).strip()   # código interno
        cleaned.append(ln)
    cleaned = glue_wrapped_lines(cleaned)

    # 3) parse linha → tail numérico e nome
    items_raw = []
    for ln in cleaned:
        toks = ln.split()
        if not toks:
            continue

        # encontra início do TAIL numérico contíguo no fim
        idx = len(toks)
        while idx > 0 and is_num_token(toks[idx-1]):
            idx -= 1
        tail = toks[idx:]
        head = toks[:idx]

        # precisamos de pelo menos 2 números no tail (qtd & valor)
        if len(tail) < 2 or not head:
            continue

        # heurística de ordem: alguns relatórios vêm [preço] [quantidade] [valor]
        # outros podem vir [quantidade] [valor] direto.
        # Tentamos interpretar os dois últimos números como (quantidade, valor) e
        # se "valor" parecer inteiro sem decimais, trocamos.
        qtd_token = tail[-2]
        valor_token = tail[-1]

        qtd = br_to_float(qtd_token)
        valor = br_to_float(valor_token)

        # Se "valor" não tiver 2 decimais e "qtd" tiver, invertemos
        def dec_places(tok):
            s = tok.replace(".", "").split(",")
            if len(s) == 2:
                return len(s[1])
            s = tok.split(".")
            if len(s) == 2:
                return len(s[1])
            return 0

        if qtd is not None and valor is not None:
            dq, dv = dec_places(qtd_token), dec_places(valor_token)
            if dv not in (2,) and dq in (2,):  # provável inversão
                qtd, valor = valor, qtd

        if qtd is None or valor is None or qtd < 0 or valor < 0:
            continue

        nome = " ".join(head).strip()
        if not re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]{3,}", nome):
            continue

        items_raw.append({"nome": nome, "quantidade": float(qtd), "valor": float(valor)})

    # 4) agrega por nome
    agg = {}
    for it in items_raw:
        k = it["nome"]
        if k not in agg:
            agg[k] = {"nome": k, "quantidade": 0.0, "valor": 0.0}
        agg[k]["quantidade"] += it["quantidade"]
        agg[k]["valor"] += it["valor"]

    # 5) ordena por valor desc
    result = sorted(agg.values(), key=lambda x: x["valor"], reverse=True)
    return result

# -------------------------
# Inputs
# -------------------------
uploaded = st.file_uploader("Envie o PDF (Curva ABC do Lince)", type=["pdf"])

default_mes = datetime.today().strftime("%m/%Y")
mes = st.text_input("Mês (ex.: 08/2025)", value=default_mes, help="Use MM/AAAA")
semana = st.text_input("Semana (ex.: 1ª semana de ago/2025)", value="", help="Como deve aparecer no Excel")

# -------------------------
# Processamento + UI
# -------------------------
if uploaded:
    all_text = extract_text_with_pypdf(uploaded)
    setor_guess = guess_setor(all_text, uploaded.name)
    setor = st.text_input("Setor", value=setor_guess)

    rows_all = parse_lince_lines_to_list(all_text)
    if not rows_all:
        st.error("Não consegui identificar linhas de produto neste PDF. Verifique se é o relatório 'Curva ABC'.")
        st.code(all_text[:2000])
        st.stop()

    # ----- Busca -----
    q = st.text_input("🔎 Buscar produto (contém):", value="").strip().upper()
    if q:
        rows = [r for r in rows_all if q in r["nome"].upper()]
    else:
        rows = rows_all[:]

    # ----- Ordenação (já vem por valor desc, mas deixo opção) -----
    order = st.selectbox("Ordenar por", ["valor (desc)", "quantidade (desc)", "nome (A→Z)"], index=0)
    if order.startswith("valor"):
        rows.sort(key=lambda x: x["valor"], reverse=True)
    elif order.startswith("quantidade"):
        rows.sort(key=lambda x: x["quantidade"], reverse=True)
    else:
        rows.sort(key=lambda x: x["nome"])

    # ----- Paginação -----
    page_size = st.selectbox("Itens por página", [20, 50, 100], index=0)
    total = len(rows)
    pages = max(1, ceil(total / page_size))
    col_p1, col_p2, col_p3 = st.columns([1, 2, 6])
    with col_p1:
        page = st.number_input("Página", min_value=1, max_value=pages, value=1, step=1)
    with col_p2:
        st.write(f"Total encontrados: **{total}** (de {len(rows_all)} detectados)")

    start = (page - 1) * page_size
    end = start + page_size
    page_rows = rows[start:end]

    # ----- Seleção (checkboxes com session_state) -----
    if "selecao" not in st.session_state:
        st.session_state.selecao = {}

    # Inicializa chaves da página atual se não existirem
    for r in page_rows:
        st.session_state.selecao.setdefault(r["nome"], True)  # pré-selecionado

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("Selecionar todos (página)"):
            for r in page_rows:
                st.session_state.selecao[r["nome"]] = True
    with c2:
        if st.button("Limpar seleção (página)"):
            for r in page_rows:
                st.session_state.selecao[r["nome"]] = False
    with c3:
        top_n = st.number_input("Pré-selecionar Top N por valor (global)", min_value=0, max_value=len(rows_all), value=0, step=1)
        if st.button("Aplicar Top N"):
            # Recria seleção: tudo False + Top N True
            st.session_state.selecao = {r["nome"]: False for r in rows_all}
            for r in rows_all[:top_n]:
                st.session_state.selecao[r["nome"]] = True

    st.markdown("---")
    # Cabeçalho da "tabela"
    hdr = st.container()
    with hdr:
        h1, h2, h3, h4 = st.columns([0.6, 4.0, 1.4, 1.4])
        h1.markdown("**Sel.**")
        h2.markdown("**Produto**")
        h3.markdown("**Quantidade**")
        h4.markdown("**Valor (R$)**")

    # Linhas da página
    box = st.container()
    for r in page_rows:
        nome = r["nome"]
        qtd = round(float(r["quantidade"]), 3)
        val = round(float(r["valor"]), 2)
        csel, cprod, cqtd, cval = box.columns([0.6, 4.0, 1.4, 1.4])
        st.session_state.selecao[nome] = csel.checkbox(
            label="",
            value=st.session_state.selecao.get(nome, True),
            key=f"chk_{nome}"
        )
        cprod.text(nome)
        cqtd.text(f"{qtd:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
        cval.text(f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    # ----- Geração do Excel -----
    st.markdown("---")
    if st.button("Gerar Excel (.xlsx)"):
        selecionados = [r for r in rows_all if st.session_state.selecao.get(r['nome'], False)]
        if not selecionados:
            st.warning("Selecione ao menos um produto.")
            st.stop()

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = workbook.add_worksheet("Produtos")

        headers = ["nome do produto", "setor", "mês", "semana", "quantidade", "valor"]
        for col, h in enumerate(headers):
            ws.write(0, col, h)

        for i, r in enumerate(selecionados, start=1):
            ws.write(i, 0, r["nome"])
            ws.write(i, 1, setor)
            ws.write(i, 2, mes)
            ws.write(i, 3, semana)
            ws.write_number(i, 4, round(float(r["quantidade"]), 3))
            ws.write_number(i, 5, round(float(r["valor"]), 2))

        # Formatação
        fmt_money = workbook.add_format({'num_format': '#,##0.00'})
        fmt_qty = workbook.add_format({'num_format': '#,##0.000'})
        ws.set_column(0, 0, 50)   # nome
        ws.set_column(1, 3, 18)   # setor/mês/semana
        ws.set_column(4, 4, 12, fmt_qty)
        ws.set_column(5, 5, 14, fmt_money)

        workbook.close()
        st.success("Excel gerado com sucesso!")
        st.download_button(
            label="⬇️ Baixar Excel",
            data=output.getvalue(),
            file_name=f"produtos_{mes.replace('/', '-')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

else:
    st.info("Envie um PDF para começar.")
