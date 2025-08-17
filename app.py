import io
import re
import os
from datetime import datetime

from pypdf import PdfReader
import streamlit as st
import xlsxwriter

# =========================
# Configuração da página
# =========================
st.set_page_config(page_title="PDF → Excel (Lite)", page_icon="🪶", layout="centered")
st.title("🪶 PDF → Excel (Lite)")
st.caption("Versão sem pandas/pdfplumber — pypdf + xlsxwriter. Seleção com checkboxes e ordenação por valor (venda).")

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

def parse_lince_lines_to_list(text: str):
    """
    Extrai itens do relatório 'Curva ABC' (Lince) de forma robusta.
    Estratégia:
      - Limpa EAN/códigos no final.
      - Separa a linha em tokens.
      - Varre da direita para a esquerda: pega os 2 últimos números como (quantidade, valor)
        e, se existir, o número anterior como preço (não usado no Excel).
      - Tudo antes vira o 'nome'.
      - Agrega por nome e ordena por 'valor' desc.
    Retorna: lista de dicts com chaves {"nome", "quantidade", "valor"}.
    """
    lines = [re.sub(r"\s{2,}", " ", (ln or "")).strip() for ln in text.splitlines()]
    lixo = (
        "Curva ABC", "Período", "CST", "ECF", "Situação Tributária",
        "Classif.", "Codigo", "CÓDIGO", "Barras", "Total do Departamento",
        "Total Geral", "www.grupotecnoweb.com.br"
    )

    items_raw = []

    for ln in lines:
        if not ln:
            continue
        if any(k in ln for k in lixo):
            continue

        # remove EAN/código no final (13 dígitos ou similares)
        ln = re.sub(r"\b\d{8,13}\b\s*$", "", ln).strip()
        # remove código interno no final (4-8 dígitos)
        ln = re.sub(r"\b\d{4,8}\b\s*$", "", ln).strip()

        tokens = ln.split()

        def is_num(tok: str) -> bool:
            # número "solto" no estilo 1.234,56 ou 1234.56
            return re.fullmatch(r"[0-9][0-9\.\,]*", tok) is not None

        nums_idx = [i for i, t in enumerate(tokens) if is_num(t)]
        if len(nums_idx) < 2:
            # precisa ter pelo menos QTD e VALOR
            continue

        # últimos dois números → (quantidade, valor)
        i_valor = nums_idx[-1]
        i_qtd = nums_idx[-2]
        valor = br_to_float(tokens[i_valor])
        qtd = br_to_float(tokens[i_qtd])

        if valor is None or qtd is None:
            continue
        if valor < 0 or qtd < 0:
            continue

        # número anterior (se existir) é possivelmente preço_unit (não usamos)
        i_preco = nums_idx[-3] if len(nums_idx) >= 3 else None

        # nome = tudo antes do preço (se existir) ou antes da quantidade
        corte = i_preco if i_preco is not None else i_qtd
        nome = " ".join(tokens[:corte]).strip()

        # sanity check do nome (precisa ter letras)
        if not re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]{3,}", nome):
            continue

        items_raw.append({"nome": nome, "quantidade": float(qtd), "valor": float(valor)})

    # agrega por nome
    agg = {}
    for it in items_raw:
        k = it["nome"]
        if k not in agg:
            agg[k] = {"nome": k, "quantidade": 0.0, "valor": 0.0}
        agg[k]["quantidade"] += it["quantidade"]
        agg[k]["valor"] += it["valor"]

    # ordena por valor desc
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
# Processamento
# -------------------------
if uploaded:
    all_text = extract_text_with_pypdf(uploaded)
    setor_guess = guess_setor(all_text, uploaded.name)
    setor = st.text_input("Setor", value=setor_guess)

    rows = parse_lince_lines_to_list(all_text)
    if not rows:
        st.error("Não consegui identificar linhas de produto neste PDF. Verifique se é o relatório 'Curva ABC'.")
        st.code(all_text[:2000])
        st.stop()

    st.subheader("Produtos detectados (ordenados por venda)")

    # -------------------------
    # UI de seleção com checkboxes (sem pandas)
    # -------------------------
    # usamos session_state para manter seleção ao interagir com os botões
    if "selecao" not in st.session_state:
        # inicia tudo como True (pré-selecionado)
        st.session_state.selecao = {r["nome"]: True for r in rows}

    # botões selecionar/limpar
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Selecionar todos"):
            for k in st.session_state.selecao.keys():
                st.session_state.selecao[k] = True
    with c2:
        if st.button("Limpar seleção"):
            for k in st.session_state.selecao.keys():
                st.session_state.selecao[k] = False

    # Cabeçalho da "tabela"
    hdr = st.container()
    with hdr:
        h1, h2, h3, h4 = st.columns([0.6, 3.4, 1.5, 1.5])
        h1.markdown("**Sel.**")
        h2.markdown("**Produto**")
        h3.markdown("**Quantidade**")
        h4.markdown("**Valor (R$)**")

    # Linhas (limitamos a altura via container para não ficar gigante)
    box = st.container()
    for r in rows:
        nome = r["nome"]
        qtd = round(float(r["quantidade"]), 3)
        val = round(float(r["valor"]), 2)
        csel, cprod, cqtd, cval = box.columns([0.6, 3.4, 1.5, 1.5])
        # checkbox com key estável
        st.session_state.selecao[nome] = csel.checkbox(
            label="",
            value=st.session_state.selecao.get(nome, True),
            key=f"chk_{nome}"
        )
        cprod.text(nome)
        cqtd.text(f"{qtd:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
        cval.text(f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    selecionados = [nome for nome, marcado in st.session_state.selecao.items() if marcado]

    st.markdown("---")
    if st.button("Gerar Excel (.xlsx)"):
        if not selecionados:
            st.warning("Selecione ao menos um produto.")
            st.stop()

        # filtra mantendo a ordem por valor desc
        final_rows = [r for r in rows if r["nome"] in selecionados]

        # cria xlsx
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = workbook.add_worksheet("Produtos")

        headers = ["nome do produto", "setor", "mês", "semana", "quantidade", "valor"]
        for col, h in enumerate(headers):
            ws.write(0, col, h)

        for i, r in enumerate(final_rows, start=1):
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
