import io
import re
import os
from datetime import datetime
from pypdf import PdfReader
import streamlit as st
import xlsxwriter

st.set_page_config(page_title="PDF → Excel (Lite)", page_icon="🪶", layout="centered")

st.title("🪶 PDF → Excel (Lite)")
st.caption("Versão sem pandas/pdfplumber: usa apenas pypdf + xlsxwriter. Ideal para Render Free.")

NUM_TOKEN = r"[0-9\.\,]+"

def br_to_float(txt: str):
    if txt is None:
        return None
    t = txt.strip()
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
        try:
            return float(t)
        except:
            pass
    t2 = t.replace(",", "")
    try:
        return float(t2)
    except:
        return None

def guess_setor(text: str, filename: str) -> str:
    m = re.search(r"Departamento:\s*([\s\S]{0,40})", text, flags=re.IGNORECASE)
    if m:
        tail = text[m.end():].splitlines()
        for ln in tail[:5]:
            t = ln.strip()
            if 2 <= len(t) <= 20 and t.isupper():
                return t
    base = os.path.basename(filename or "")
    base_up = base.upper()
    for chave in ["FRIOS", "AÇOUGUE", "PADARIA", "HORTIFRUTI", "BEBIDAS", "MERCEARIA"]:
        if chave in base_up:
            start = base_up.find(chave)
            end = min(len(base_up), start + len(chave) + 2)
            return re.sub(r"[^A-Z0-9]", "", base_up[start:end])
    return "N/D"

def extract_text_with_pypdf(file):
    reader = PdfReader(file)
    texts = []
    for page in reader.pages:
        try:
            texts.append(page.extract_text() or "")
        except Exception:
            texts.append("")
    return "\n".join(texts)

def parse_lince_lines_to_list(text: str):
    items = []
    lines = [re.sub(r"\s{2,}", " ", ln).strip() for ln in text.splitlines()]
    lixo = (
        "Curva ABC", "Período", "CST", "ECF", "Situação Tributária",
        "Classif.", "Codigo", "Barras", "Total do Departamento",
        "Total Geral", "www.grupotecnoweb.com.br"
    )

    patt = re.compile(
        rf"^(?P<nome>.+?)\s+(?P<preco>{NUM_TOKEN})\s+(?P<qtd>{NUM_TOKEN})\s+(?P<valor>{NUM_TOKEN})(\s+.+)?$"
    )

    for ln in lines:
        if not ln or any(k in ln for k in lixo):
            continue
        ln_clean = re.sub(r"\b\d{8,13}\b$", "", ln).strip()
        ln_clean = re.sub(r"\b\d{4,8}\b\s*$", "", ln_clean).strip()
        m = patt.match(ln_clean)
        if not m:
            continue
        nome = m.group("nome").strip()
        preco = br_to_float(m.group("preco"))
        qtd   = br_to_float(m.group("qtd"))
        val   = br_to_float(m.group("valor"))
        if preco is None or preco <= 0: 
            continue
        if val is None or val < 0: 
            continue
        if qtd is None or qtd < 0: 
            continue
        if not re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]{3,}", nome):
            continue
        items.append({"nome": nome, "quantidade": qtd, "valor": val})

    agg = {}
    for it in items:
        key = it["nome"]
        if key not in agg:
            agg[key] = {"nome": key, "quantidade": 0.0, "valor": 0.0}
        agg[key]["quantidade"] += float(it["quantidade"] or 0.0)
        agg[key]["valor"] += float(it["valor"] or 0.0)

    result = sorted(agg.values(), key=lambda x: x["valor"], reverse=True)
    return result

uploaded = st.file_uploader("Envie o PDF (Curva ABC do Lince)", type=["pdf"])
default_mes = datetime.today().strftime("%m/%Y")
mes = st.text_input("Mês (ex.: 08/2025)", value=default_mes, help="Use MM/AAAA")
semana = st.text_input("Semana (ex.: 1ª semana de ago/2025)", value="", help="Como deve aparecer no Excel")

if uploaded:
    all_text = extract_text_with_pypdf(uploaded)
    setor_guess = guess_setor(all_text, uploaded.name)
    setor = st.text_input("Setor", value=setor_guess)

    rows = parse_lince_lines_to_list(all_text)
    if not rows:
        st.error("Não consegui identificar linhas de produto neste PDF. Verifique se é o relatório 'Curva ABC'.")
        st.code(all_text[:2000])
        st.stop()

    st.subheader("Produtos detectados")
    st.write([{"nome": r["nome"], "quantidade": round(r["quantidade"],3), "valor": round(r["valor"],2)} for r in rows][:50])

    nomes = [r["nome"] for r in rows]
    selecionados = st.multiselect("Selecione os produtos para o Excel", options=nomes, default=nomes[: min(10, len(nomes))])

    if st.button("Gerar Excel (.xlsx)"):
        if not selecionados:
            st.warning("Selecione ao menos um produto.")
            st.stop()

        final_rows = [r for r in rows if r["nome"] in selecionados]

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

        fmt_money = workbook.add_format({'num_format': '#,##0.00'})
        fmt_qty = workbook.add_format({'num_format': '#,##0.000'})
        ws.set_column(0, 0, 50)
        ws.set_column(1, 3, 18)
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
