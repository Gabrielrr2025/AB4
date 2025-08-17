import io
import re
import os
from math import ceil
from datetime import datetime

from pypdf import PdfReader
import streamlit as st
import xlsxwriter

# =========================
# Config
# =========================
st.set_page_config(page_title="PDF → Excel - Shopping do Pão", page_icon="🛍️", layout="wide")
st.title("🛍️ PDF → Excel - Shopping do Pão")
st.caption("Parser robusto para relatórios Curva ABC (Lince) - Nomes limpos, busca, paginação, seleção múltipla.")

# Lista de setores válidos
SETORES_VALIDOS = [
    "Padaria", 
    "Frios", 
    "Restaurante", 
    "Confeitaria Fina", 
    "Confeitaria Trad", 
    "Salgados", 
    "Lanchonete"
]

# -------------------------
# Utilidades
# -------------------------
def br_to_float(txt: str):
    """Converte '1.234,56' → 1234.56; e '1,234.56' → 1234.56."""
    if not txt:
        return None
    t = txt.strip()
    if not t:
        return None
    if "," in t:
        t1 = t.replace(".", "").replace(",", ".")
        try:
            return float(t1)
        except Exception:
            pass
    t2 = t.replace(",", "")
    try:
        return float(t2)
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
    """Extrai texto de todas as páginas (tolerante a erros)."""
    reader = PdfReader(file)
    texts = []
    for page in reader.pages:
        try:
            texts.append(page.extract_text() or "")
        except Exception:
            texts.append("")
    return "\n".join(texts)

def guess_setor(text: str, filename: str) -> str:
    """Detecta setor automaticamente usando a lista de setores válidos."""
    
    # Busca no texto após "Departamento:"
    m = re.search(r"Departamento:\s*([\s\S]{0,100})", text, flags=re.IGNORECASE)
    if m:
        departamento_section = m.group(1)
        # Verifica setores válidos na seção
        for setor in SETORES_VALIDOS:
            if setor.upper() in departamento_section.upper():
                return setor
        
        # Busca por números seguidos de nomes de setores
        tail = text[m.end():].splitlines()
        for ln in tail[:5]:
            t = (ln or "").strip()
            if 2 <= len(t) <= 30:
                for setor in SETORES_VALIDOS:
                    if setor.upper() in t.upper():
                        return setor
    
    # Busca no texto completo
    text_upper = text.upper()
    for setor in SETORES_VALIDOS:
        if setor.upper() in text_upper:
            return setor
    
    # Busca no nome do arquivo
    if filename:
        base_up = os.path.basename(filename).upper()
        for setor in SETORES_VALIDOS:
            if setor.upper() in base_up:
                return setor
        
        # Mapeamento de palavras-chave
        mapeamento = {
            "FRIOS": "Frios",
            "AÇOUGUE": "Frios",
            "ACOUGUE": "Frios", 
            "PADARIA": "Padaria",
            "CONFEIT": "Confeitaria Fina",
            "DOCE": "Confeitaria Trad",
            "SALGADO": "Salgados",
            "LANCHE": "Lanchonete",
            "RESTAUR": "Restaurante"
        }
        
        for keyword, setor in mapeamento.items():
            if keyword in base_up:
                return setor
    
    return "Lanchonete"  # padrão baseado no exemplo

def glue_wrapped_lines(lines):
    """
    Une linhas quebradas: se a linha atual não tem tail numérico (>=2 tokens numéricos no fim)
    e a próxima é majoritariamente numérica, concatena.
    """
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
    """
    Remove:
      - EAN/GTIN (12+ dígitos) em qualquer posição
      - 1º token numérico curto no início (código do item, ex.: 4051)
    Mantém tokens 'mistos' de nome (ex.: '200ML', 'UN', 'KG') porque têm letras.
    """
    out = []
    removed_leading_code = False
    for idx, t in enumerate(tokens):
        if re.fullmatch(r"\d{12,}", t):  # EAN em qualquer lugar
            continue
        if not removed_leading_code and idx == 0 and re.fullmatch(r"\d{3,6}", t):
            removed_leading_code = True
            continue
        out.append(t)
    return out

def parse_lince_lines_to_list(text: str):
    """
    Extrai itens do relatório Curva ABC (Lince):
      - normaliza, remove cabeçalhos/rodapés;
      - cola linhas quebradas;
      - limpa EAN/código do começo;
      - identifica tail numérico no fim e pega (quantidade, valor) corretos;
      - nome = head textual (sem números do tail).
    Retorna: lista de dicts {"nome","quantidade","valor"} ordenada por 'valor' desc.
    """
    # 1) normaliza e remove cabeçalhos/rodapés
    lines = [re.sub(r"\s{2,}", " ", (ln or "")).strip() for ln in text.splitlines()]
    lixo = ("Curva ABC","Período","CST","ECF","Situação Tributária","Classif.","Codigo","CÓDIGO",
            "Barras","Total do Departamento","Total Geral","www.grupotecnoweb.com.br")
    lines = [ln for ln in lines if ln and not any(k in ln for k in lixo)]

    # 2) remove EAN/código no final; cola linhas
    cleaned = []
    for ln in lines:
        ln = re.sub(r"\b\d{8,13}\b\s*$", "", ln).strip()   # EAN no final
        ln = re.sub(r"\b\d{4,8}\b\s*$", "", ln).strip()    # código no final
        cleaned.append(ln)
    cleaned = glue_wrapped_lines(cleaned)

    items_raw = []
    for ln in cleaned:
        toks = ln.split()
        if not toks:
            continue
        toks = clean_tokens(toks)  # remove código inicial e EANs internos
        if not toks:
            continue

        # acha início do tail numérico contíguo no fim
        idx = len(toks)
        while idx > 0 and is_num_token(toks[idx-1]):
            idx -= 1
        head = toks[:idx]
        tail = toks[idx:]

        if len(tail) < 2 or not head:
            continue

        # Candidato a VALOR: token com 2 casas decimais mais à direita
        cand_valores = [i for i, t in enumerate(tail) if dec_places(t) == 2 and br_to_float(t) is not None]
        if cand_valores:
            i_val = cand_valores[-1]
            valor = br_to_float(tail[i_val])
            # QTD: procurar à esquerda do valor um token com 3 casas decimais; senão, pegar o válido mais próximo
            i_qtd = None
            for j in range(i_val - 1, -1, -1):
                if dec_places(tail[j]) == 3 and br_to_float(tail[j]) is not None:
                    i_qtd = j
                    break
            if i_qtd is None:
                for j in range(i_val - 1, -1, -1):
                    if br_to_float(tail[j]) is not None:
                        i_qtd = j
                        break
            if i_qtd is None or valor is None:
                continue
            qtd = br_to_float(tail[i_qtd])
        else:
            # fallback: usa os 2 últimos números como qtd/valor
            qtd = br_to_float(tail[-2])
            valor = br_to_float(tail[-1])

        if qtd is None or valor is None or qtd < 0 or valor < 0:
            continue

        # Nome: somente head textual; filtra "números soltos" que eventualmente ficaram no head
        # (mantém tokens com letras, como 200ML/UN/KG)
        head_clean = []
        for t in head:
            if is_num_token(t):
                continue  # elimina número puro perdido no head
            head_clean.append(t)
        nome = " ".join(head_clean).strip()

        # sanity: precisa ter letras
        if not re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", nome):
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
    return sorted(agg.values(), key=lambda x: x["valor"], reverse=True)

# -------------------------
# Interface Principal
# -------------------------
uploaded = st.file_uploader("📁 Envie o PDF (Curva ABC do Lince)", type=["pdf"])

# Interface melhorada para configurações
st.subheader("⚙️ Configurações do Relatório")

col1, col2, col3 = st.columns(3)

with col1:
    # Mês dropdown
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    mes_atual = datetime.now().month - 1
    mes_selecionado = st.selectbox("📅 Mês", options=meses, index=mes_atual)

with col2:
    # Semana como número
    semana = st.selectbox("📊 Semana", options=[1, 2, 3, 4], index=0)

with col3:
    # Setor será preenchido após upload do PDF
    setor_placeholder = st.empty()

# -------------------------
# UI + Geração
# -------------------------
if uploaded:
    # Extrai texto do PDF
    with st.spinner("🔄 Processando PDF..."):
        all_text = extract_text_with_pypdf(uploaded)
    
    # Detecta setor automaticamente
    setor_guess = guess_setor(all_text, uploaded.name)
    
    # Dropdown para setor
    with col3:
        try:
            setor_index = SETORES_VALIDOS.index(setor_guess)
        except ValueError:
            setor_index = 6  # Lanchonete como padrão
        setor = st.selectbox("🏪 Setor", options=SETORES_VALIDOS, index=setor_index)

    # Parse dos produtos
    with st.spinner("🔍 Analisando produtos..."):
        rows_all = parse_lince_lines_to_list(all_text)
    
    if not rows_all:
        st.error("❌ Não consegui identificar produtos neste PDF.")
        with st.expander("🔍 Ver texto extraído (debug)"):
            st.code(all_text[:2000])
        st.stop()

    st.success(f"✅ {len(rows_all)} produtos detectados!")

    # ----- Controles de Busca e Filtros -----
    st.subheader("🔍 Busca e Filtros")
    
    col_search, col_order = st.columns([2, 1])
    with col_search:
        q = st.text_input("🔎 Buscar produto (contém):", value="").strip().upper()
    with col_order:
        order = st.selectbox("📊 Ordenar por", ["valor (desc)", "quantidade (desc)", "nome (A→Z)"], index=0)

    # Aplica busca
    if q:
        rows = [r for r in rows_all if q in r["nome"].upper()]
        if not rows:
            st.warning(f"Nenhum produto encontrado com '{q}'")
    else:
        rows = rows_all[:]

    # Aplica ordenação
    if order.startswith("valor"):
        rows.sort(key=lambda x: x["valor"], reverse=True)
    elif order.startswith("quantidade"):
        rows.sort(key=lambda x: x["quantidade"], reverse=True)
    else:
        rows.sort(key=lambda x: x["nome"])

    # ----- Paginação -----
    col_page1, col_page2, col_page3 = st.columns([1, 1, 2])
    with col_page1:
        page_size = st.selectbox("📄 Itens por página", [20, 50, 100], index=0)
    
    total = len(rows)
    pages = max(1, ceil(total / page_size))
    
    with col_page2:
        page = st.number_input("Página", min_value=1, max_value=pages, value=1, step=1)
    with col_page3:
        st.info(f"📊 **{total}** produtos encontrados (de {len(rows_all)} detectados)")

    start = (page - 1) * page_size
    end = start + page_size
    page_rows = rows[start:end]

    # ----- Seleção (checkboxes com session_state) -----
    st.subheader("✅ Seleção de Produtos")
    
    if "selecao" not in st.session_state:
        st.session_state.selecao = {}

    # Inicializa chaves de todos os produtos se não existirem
    for r in rows_all:
        st.session_state.selecao.setdefault(r["nome"], True)  # pré-selecionado

    # Controles de seleção
    col_sel1, col_sel2, col_sel3, col_sel4 = st.columns(4)
    with col_sel1:
        if st.button("✅ Selecionar todos (página)"):
            for r in page_rows:
                st.session_state.selecao[r["nome"]] = True
    with col_sel2:
        if st.button("❌ Limpar seleção (página)"):
            for r in page_rows:
                st.session_state.selecao[r["nome"]] = False
    with col_sel3:
        top_n = st.number_input("🎯 Top N por valor (global)", min_value=0, max_value=len(rows_all), value=10, step=1)
    with col_sel4:
        if st.button("🎯 Aplicar Top N"):
            # Recria seleção: tudo False + Top N True
            st.session_state.selecao = {r["nome"]: False for r in rows_all}
            for r in rows_all[:top_n]:
                st.session_state.selecao[r["nome"]] = True

    # Info sobre seleção atual
    selecionados_count = sum(1 for r in rows_all if st.session_state.selecao.get(r['nome'], False))
    valor_selecionado = sum(r['valor'] for r in rows_all if st.session_state.selecao.get(r['nome'], False))
    
    if selecionados_count > 0:
        st.info(f"📊 **{selecionados_count}** produtos selecionados | Valor total: **R$ {valor_selecionado:,.2f}**".replace(',', 'X').replace('.', ',').replace('X', '.'))

    # ----- Tabela de Produtos -----
    st.markdown("---")
    
    # Cabeçalho da "tabela"
    hdr = st.container()
    with hdr:
        h1, h2, h3, h4 = st.columns([0.6, 4.5, 1.4, 1.5])
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
        csel, cprod, cqtd, cval = box.columns([0.6, 4.5, 1.4, 1.5])
        st.session_state.selecao[nome] = csel.checkbox(
            label="",
            value=st.session_state.selecao.get(nome, True),
            key=f"chk_{nome}_{page}"  # inclui página para evitar conflitos
        )
        cprod.text(nome)  # <<<<<< NOME LIMPO, sem números extras
        cqtd.text(f"{qtd:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
        cval.text(f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    # ----- Geração do Excel -----
    st.markdown("---")
    st.subheader("📊 Gerar Excel")
    
    col_excel1, col_excel2 = st.columns([3, 1])
    
    with col_excel2:
        st.write("") # espaço
        if st.button("📊 **Gerar Excel**", type="primary", use_container_width=True):
            selecionados = [r for r in rows_all if st.session_state.selecao.get(r['nome'], False)]
            if not selecionados:
                st.warning("⚠️ Selecione pelo menos um produto.")
                st.stop()

            # Prepara arquivo Excel
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            ws = workbook.add_worksheet("Produtos")

            # Cabeçalhos conforme especificação
            headers = ["nome do produto", "setor", "mês", "semana", "quantidade", "valor"]
            for col, h in enumerate(headers):
                ws.write(0, col, h)

            # Dados dos produtos selecionados
            for i, r in enumerate(selecionados, start=1):
                ws.write(i, 0, r["nome"])
                ws.write(i, 1, setor)
                ws.write(i, 2, mes_selecionado)
                ws.write(i, 3, semana)
                ws.write_number(i, 4, round(float(r["quantidade"]), 3))
                ws.write_number(i, 5, round(float(r["valor"]), 2))

            # Formatação
            fmt_money = workbook.add_format({'num_format': 'R$ #,##0.00'})
            fmt_qty = workbook.add_format({'num_format': '#,##0.000'})
            
            ws.set_column(0, 0, 50)   # nome do produto
            ws.set_column(1, 1, 18)   # setor
            ws.set_column(2, 2, 15)   # mês
            ws.set_column(3, 3, 10)   # semana
            ws.set_column(4, 4, 15, fmt_qty)   # quantidade
            ws.set_column(5, 5, 15, fmt_money) # valor

            workbook.close()
            
            # Nome do arquivo conforme especificação
            nome_arquivo = f"produtos_{setor.lower().replace(' ', '_')}_{mes_selecionado.lower()}_semana{semana}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            
            st.success("✅ Excel gerado com sucesso!")
            st.download_button(
                label="⬇️ Baixar Excel",
                data=output.getvalue(),
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with col_excel1:
        if selecionados_count > 0:
            st.write("**Preview do Excel:**")
            preview_df = []
            count = 0
            for r in rows_all:
                if st.session_state.selecao.get(r['nome'], False) and count < 5:
                    preview_df.append({
                        "nome do produto": r["nome"],
                        "setor": setor,
                        "mês": mes_selecionado,
                        "semana": semana,
                        "quantidade": round(r["quantidade"], 3),
                        "valor": round(r["valor"], 2)
                    })
                    count += 1
            
            if preview_df:
                st.table(preview_df)
                if len(preview_df) < selecionados_count:
                    st.caption(f"... e mais {selecionados_count - len(preview_df)} produtos")

else:
    # Tela inicial com instruções
    st.info("📋 **Como usar:**")
    st.markdown("""
    1. 📄 **Faça upload** do PDF com relatório 'Curva ABC' do sistema Lince
    2. ⚙️ **Configure** o mês, semana e setor (detectado automaticamente)
    3. 🔍 **Use a busca** para encontrar produtos específicos
    4. ✅ **Selecione** os produtos que deseja exportar (use Top N para facilitar)
    5. 📊 **Gere** o arquivo Excel com as colunas especificadas
    
    **Colunas do Excel:** nome do produto, setor, mês, semana, quantidade, valor
    """)
    
    # Exemplo visual dos setores e formato do arquivo
    st.markdown("---")
    col_info1, col_info2 = st.columns(2)
    
    with col_info1:
        st.markdown("### 🏪 Setores Disponíveis:")
        for setor in SETORES_VALIDOS:
            st.markdown(f"• {setor}")
    
    with col_info2:
        st.markdown("### 📊 Exemplo de saída:")
        exemplo_df = [
            {
                "nome do produto": "CAFE E OU CAFE C/LEITE 200ML",
                "setor": "Lanchonete", 
                "mês": "Agosto",
                "semana": 1,
                "quantidade": 506.000,
                "valor": 3491.40
            },
            {
                "nome do produto": "PAO C QUEIJO MINAS",
                "setor": "Lanchonete",
                "mês": "Agosto", 
                "semana": 1,
                "quantidade": 174.000,
                "valor": 2697.00
            }
        ]
        st.table(exemplo_df)
