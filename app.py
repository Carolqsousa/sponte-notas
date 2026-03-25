"""
Cultura Inglesa — Mapa de Notas
App Streamlit com suporte a nova planilha e atualização de planilha existente.
"""

import streamlit as st
import requests
import pandas as pd
import time
import io
import base64
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ─── Configuração das unidades ─────────────────────────────────────────────────
UNIDADES = {
    "Young":      {"token": st.secrets["YOUNG_TOKEN"],      "codigo": "71976", "semestre": "2026.1"},
    "Boa Viagem": {"token": st.secrets["BOA_VIAGEM_TOKEN"], "codigo": "71961", "semestre": "2026.1"},
    "Setúbal":    {"token": st.secrets["SETUBAL_TOKEN"],    "codigo": "71977", "semestre": "2026.1"},
    "Natal":      {"token": st.secrets["NATAL_TOKEN"],      "codigo": "71978", "semestre": "2026/01"},
}

BASE_URL  = "https://webservices.sponteweb.com.br/WSApiSponteRest/api"
SITUATION = 1

# ─── Estrutura de notas ────────────────────────────────────────────────────────
NOME_CURTO_A = {
    "Speaking / Oral (PC)":                    "Speaking/Oral",
    "Language / Gramática & Vocabulário (PC)": "Gramática & Vocab",
    "Homework / Tarefa de casa":               "Homework",
    "Reading / Leitura":                       "Reading",
    "Language / Gramática & Vocabulário":      "Gramática & Vocab",
    "Listening / Compreensão Auditiva":        "Listening",
    "Writing / Escrita":                       "Writing",
    "Speaking / Oral":                         "Speaking/Oral",
    "Speaking / Oral ":                        "Speaking/Oral",
}
COMP_PC        = ["Speaking / Oral (PC)", "Language / Gramática & Vocabulário (PC)", "Homework / Tarefa de casa"]
COMP_MID_FINAL = ["Homework / Tarefa de casa", "Reading / Leitura", "Language / Gramática & Vocabulário",
                  "Listening / Compreensão Auditiva", "Writing / Escrita", "Speaking / Oral", "Speaking / Oral "]
PROVAS_A = {
    "Progress Check": ("PC",    COMP_PC),
    "Mid-term":       ("Mid",   COMP_MID_FINAL),
    "Final":          ("Final", COMP_MID_FINAL),
}
PROVAS_B_SIMPLES = {
    "AVALIAÇÃO 1": ("Av1", "Reading And Use of English"),
    "AVALIAÇÃO 2": ("Av2", "Writing"),
    "AVALIAÇÃO 3": ("Av3", "Listening"),
    "AVALIAÇÃO 4": ("Av4", "Homework"),
}
PHASES_FORMATO_B = {
    "ADVANCED 1 (F)", "ADVANCED 2 (F)",
    "MASTERY 1 (F)", "MASTERY 2 (F)",
    "VANTAGE 1 (F)", "VANTAGE 2 (F)",
    "UPPER INTERMEDIATE 3 (F)",
}
COLUNAS_EXTRA = ["Status", "Action Plan", "Coordinator's Comment"]

def build_colunas_a():
    cols = []
    for prova, (pre, comps) in PROVAS_A.items():
        seen = set()
        for c in comps:
            col = f"{pre} - {NOME_CURTO_A[c]}"
            if col not in seen:
                cols.append(col)
                seen.add(col)
        cols.append(f"Média {pre}")
    return cols

COLUNAS_A     = build_colunas_a()
COLUNAS_B     = [f"{pre} - {comp}" for _, (pre, comp) in PROVAS_B_SIMPLES.items()] + ["Média Geral Av"]
COLUNAS_INFO_A = ["Turma", "Phase", "Professor", "student_id", "Nome Aluno", "Situação", "Média Geral"]
COLUNAS_INFO_B = ["Turma", "Phase", "Professor", "student_id", "Nome Aluno", "Situação"]
TODAS_COLS_A  = COLUNAS_INFO_A + COLUNAS_A + COLUNAS_EXTRA
TODAS_COLS_B  = COLUNAS_INFO_B + COLUNAS_B + COLUNAS_EXTRA

# ─── Funções API ───────────────────────────────────────────────────────────────
def headers(token):
    return {"Accept": "application/json", "Content-Type": "application/json", "api_key": token}

def get_phases_map(token):
    r = requests.get(f"{BASE_URL}/phases", headers=headers(token))
    return {p["name"]: p["phase_id"] for p in r.json()} if r.ok else {}

def get_turmas(token, semestre):
    r = requests.get(f"{BASE_URL}/classes", headers=headers(token))
    if not r.ok:
        return []
    return [t for t in r.json() if t.get("situation") == SITUATION and semestre in t.get("name", "")]

def get_detalhes(token, class_id):
    r = requests.post(f"{BASE_URL}/classes", headers=headers(token), json={"class_id": class_id})
    return r.json() if r.ok else None

def get_nome_aluno(token, student_id):
    r = requests.post(f"{BASE_URL}/students", headers=headers(token), json={"student_id": student_id})
    if not r.ok:
        return ""
    d = r.json()
    if isinstance(d, list):
        d = d[0] if d else {}
    return d.get("name") or ""

def get_notas(token, student_id, class_id, phase_id):
    r = requests.post(f"{BASE_URL}/scores", headers=headers(token),
                      json={"student_id": student_id, "class_id": class_id, "phase_id": phase_id})
    if not r.ok:
        return None
    d = r.json()
    if d.get("errors") is False and not d.get("grades"):
        return None
    return d

def extrair_notas_a(notas_data):
    resultado = {}
    if not notas_data:
        return resultado
    grades = notas_data.get("grades", [])
    por_prova = {}
    for g in grades:
        if not g.get("score"):
            continue
        prova       = g.get("test_name", "").strip()
        comp        = g.get("evaluation_name", "").strip()
        nota        = str(g.get("score", "")).replace(".", ",")
        weight      = g.get("weight", 1) or 1
        test_weight = g.get("test_weight", 1) or 1
        por_prova.setdefault(prova, []).append((comp, nota, weight, test_weight))

    medias_provas = {}
    for prova, items in por_prova.items():
        entry = next(((pre, comps) for p, (pre, comps) in PROVAS_A.items()
                      if p.upper() == prova.upper()), None)
        if not entry:
            continue
        pre, _ = entry
        for comp, nota, weight, tw in items:
            nome = NOME_CURTO_A.get(comp, NOME_CURTO_A.get(comp.strip(), comp))
            col  = f"{pre} - {nome}"
            if col in COLUNAS_A:
                resultado[col] = nota
        soma_n, soma_p, tw_val = 0.0, 0.0, 1
        for comp, nota, weight, tw in items:
            try:
                soma_n += float(nota.replace(",", ".")) * weight
                soma_p += weight
                tw_val  = tw
            except:
                pass
        if soma_p > 0:
            media = round(soma_n / soma_p, 2)
            resultado[f"Média {pre}"] = str(media).replace(".", ",")
            medias_provas[pre] = (media, tw_val)

    if medias_provas:
        soma_m  = sum(m * tw for m, tw in medias_provas.values())
        soma_tw = sum(tw for _, tw in medias_provas.values())
        resultado["Média Geral"] = str(round(soma_m / soma_tw, 2)).replace(".", ",")
    return resultado

def extrair_notas_b(notas_data):
    resultado = {}
    if not notas_data:
        return resultado
    grades = notas_data.get("grades", [])
    notas_av = {}
    for g in grades:
        if not g.get("score"):
            continue
        prova = g.get("test_name", "").strip().upper()
        comp  = g.get("evaluation_name", "").strip()
        nota  = str(g.get("score", "")).replace(".", ",")
        notas_av.setdefault(prova, {})[comp] = nota

    vals = []
    for prova_key, (pre, comp_esp) in PROVAS_B_SIMPLES.items():
        nota = notas_av.get(prova_key, {}).get(comp_esp, "")
        col  = f"{pre} - {comp_esp}"
        resultado[col] = nota
        if nota:
            try:
                vals.append(float(nota.replace(",", ".")))
            except:
                pass
    if vals:
        resultado["Média Geral Av"] = str(round(sum(vals)/len(vals), 2)).replace(".", ",")
    return resultado

# ─── Mescla com planilha existente ────────────────────────────────────────────
def mesclar_com_existente(novos_dados_por_prof, arquivo_existente):
    """
    Lê planilha existente, preserva Action Plan, Coordinator's Comment.
    Alunos que saíram da turma ficam com Status = 'Não consta mais na turma'.
    """
    wb_old = load_workbook(arquivo_existente)
    comentarios = {}  # {(sheet_name, student_id): {Action Plan, Coordinator's Comment}}

    for sheet_name in wb_old.sheetnames:
        ws = wb_old[sheet_name]
        headers_row = [c.value for c in ws[1]]
        try:
            idx_sid  = headers_row.index("student_id")
            idx_ap   = headers_row.index("Action Plan") if "Action Plan" in headers_row else None
            idx_cc   = headers_row.index("Coordinator's Comment") if "Coordinator's Comment" in headers_row else None
        except:
            continue

        for row in ws.iter_rows(min_row=2, values_only=True):
            sid = row[idx_sid] if idx_sid < len(row) else None
            if not sid:
                continue
            comentarios[(sheet_name, int(sid))] = {
                "Action Plan":           row[idx_ap] if idx_ap and idx_ap < len(row) else "",
                "Coordinator's Comment": row[idx_cc] if idx_cc and idx_cc < len(row) else "",
            }

    # IDs ativos por aba
    ids_ativos_por_aba = {}
    for prof, fmts in novos_dados_por_prof.items():
        for fmt, linhas in fmts.items():
            if not linhas:
                continue
            sufixo   = f" - {fmt}" if fmts["A"] and fmts["B"] else ""
            nome_aba = (prof + sufixo)[:31]
            for ch in ['/', '\\', '*', '?', ':', '[', ']']:
                nome_aba = nome_aba.replace(ch, '-')
            ids_ativos_por_aba[nome_aba] = {l["student_id"] for l in linhas}

    # Aplica comentários e status
    for prof, fmts in novos_dados_por_prof.items():
        for fmt, linhas in fmts.items():
            sufixo   = f" - {fmt}" if fmts["A"] and fmts["B"] else ""
            nome_aba = (prof + sufixo)[:31]
            for ch in ['/', '\\', '*', '?', ':', '[', ']']:
                nome_aba = nome_aba.replace(ch, '-')

            for linha in linhas:
                sid  = linha["student_id"]
                chave = (nome_aba, int(sid))
                if chave in comentarios:
                    linha["Action Plan"]           = comentarios[chave]["Action Plan"] or ""
                    linha["Coordinator's Comment"] = comentarios[chave]["Coordinator's Comment"] or ""

            # Alunos que saíram
            ids_novos = {l["student_id"] for l in linhas}
            todas_cols = TODAS_COLS_A if fmt == "A" else TODAS_COLS_B
            for (sn, sid), coment in comentarios.items():
                if sn == nome_aba and sid not in ids_novos:
                    linha_saiu = {col: "" for col in todas_cols}
                    linha_saiu["student_id"]              = sid
                    linha_saiu["Status"]                  = "⚠️ Não consta mais na turma"
                    linha_saiu["Action Plan"]             = coment["Action Plan"] or ""
                    linha_saiu["Coordinator's Comment"]   = coment["Coordinator's Comment"] or ""
                    linhas.append(linha_saiu)

    return novos_dados_por_prof

# ─── Formatação Excel ──────────────────────────────────────────────────────────
FILL_HEADER  = PatternFill("solid", fgColor="1A2B6B")
FILL_RED     = PatternFill("solid", fgColor="FFCCCC")
FILL_YELLOW  = PatternFill("solid", fgColor="FFFACD")
FILL_GRAY    = PatternFill("solid", fgColor="EEEEEE")
FONT_WHITE   = Font(bold=True, color="FFFFFF")

def is_nota_col(col):
    return any(col.startswith(p) for p in ("PC -","Mid -","Final -","Av1 -","Av2 -","Av3 -","Av4 -","Média"))

def parse_nota(val):
    try:
        return float(str(val).replace(",", "."))
    except:
        return None

def formatar_aba(ws, colunas):
    for cell in ws[1]:
        cell.fill = FILL_HEADER
        cell.font = FONT_WHITE
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    nota_cols   = {i+1 for i, c in enumerate(colunas) if is_nota_col(c)}
    status_col  = next((i+1 for i, c in enumerate(colunas) if c == "Status"), None)

    for row in ws.iter_rows(min_row=2):
        # Linhas de alunos que saíram → cinza
        status_val = row[status_col-1].value if status_col else ""
        saiu = "Não consta" in str(status_val or "")

        for cell in row:
            if saiu:
                cell.fill = FILL_GRAY
                continue
            if cell.column not in nota_cols:
                continue
            n = parse_nota(cell.value)
            if n is None:
                cell.fill = FILL_YELLOW
            elif n < 7:
                cell.fill = FILL_RED

    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 2, 35)

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "H2"
    for i, col in enumerate(colunas, start=1):
        if col in ("Phase", "student_id"):
            ws.column_dimensions[get_column_letter(i)].hidden = True

# ─── Geração da planilha ───────────────────────────────────────────────────────
def coletar_dados(token, semestre, progress_bar, status_text):
    phases_map = get_phases_map(token)
    turmas     = get_turmas(token, semestre)
    if not turmas:
        return None, "Nenhuma turma aberta encontrada."

    dados       = {}
    cache_nomes = {}
    total       = len(turmas)

    for i, turma in enumerate(turmas):
        class_id   = turma["class_id"]
        nome_turma = turma["name"]
        status_text.text(f"⏳ Turma {i+1}/{total}: {nome_turma}")
        progress_bar.progress(i / total)

        detalhes = get_detalhes(token, class_id)
        if not detalhes:
            continue

        professor  = detalhes.get("professor_name") or "Sem Professor"
        alunos     = detalhes.get("members", [])
        phase_nome = next((s["phase"] for s in detalhes.get("schedule", []) if s.get("phase")), None)
        phase_id   = phases_map.get(phase_nome) if phase_nome else None
        formato_b  = phase_nome in PHASES_FORMATO_B if phase_nome else False
        fmt        = "B" if formato_b else "A"
        todas_cols = TODAS_COLS_B if formato_b else TODAS_COLS_A

        if professor not in dados:
            dados[professor] = {"A": [], "B": []}

        for aluno in alunos:
            sid = aluno.get("student_id")
            if not sid:
                continue
            if sid not in cache_nomes:
                cache_nomes[sid] = get_nome_aluno(token, sid)
                time.sleep(0.05)

            notas_dict = {}
            situacao   = ""
            if phase_id:
                nd = get_notas(token, sid, class_id, phase_id)
                time.sleep(0.07)
                if nd:
                    situacao   = nd.get("situation", "") or ""
                    notas_dict = extrair_notas_b(nd) if formato_b else extrair_notas_a(nd)

            linha = {
                "Turma": nome_turma, "Phase": phase_nome or "",
                "Professor": professor, "student_id": sid,
                "Nome Aluno": cache_nomes[sid], "Situação": situacao,
                "Status": "",
            }
            linha.update(notas_dict)
            for col in todas_cols:
                linha.setdefault(col, "")
            dados[professor][fmt].append(linha)

    progress_bar.progress(1.0)
    return dados, None

def exportar_excel(dados):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for professor, fmts in sorted(dados.items()):
            for fmt, linhas in fmts.items():
                if not linhas:
                    continue
                todas_cols = TODAS_COLS_A if fmt == "A" else TODAS_COLS_B
                df         = pd.DataFrame(linhas, columns=todas_cols)
                sufixo     = f" - {fmt}" if fmts["A"] and fmts["B"] else ""
                nome_aba   = (professor + sufixo)[:31]
                for ch in ['/', '\\', '*', '?', ':', '[', ']']:
                    nome_aba = nome_aba.replace(ch, '-')
                df.to_excel(writer, sheet_name=nome_aba, index=False)
                formatar_aba(writer.sheets[nome_aba], todas_cols)
    output.seek(0)
    return output

# ─── Interface ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Cultura Inglesa — Mapa de Notas",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Barlow', sans-serif; background-color: #f4f6fb; }

    .header {
        background: white;
        border-bottom: 3px solid #e32119;
        padding: 1.2rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        display: flex;
        align-items: center;
        gap: 1.5rem;
        box-shadow: 0 2px 12px rgba(26,43,107,0.08);
    }
    .header-text h1 {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 1.8rem;
        font-weight: 700;
        margin: 0;
        color: #1a2b6b;
        letter-spacing: 0.3px;
    }
    .header-text p { margin: 0.1rem 0 0; color: #888; font-size: 0.85rem; }
    .header img { height: 48px; }

    .unit-card {
        background: white;
        border: 2px solid #e8edf8;
        border-radius: 14px;
        padding: 1.5rem 1rem 1rem;
        text-align: center;
        box-shadow: 0 2px 8px rgba(26,43,107,0.06);
    }
    .unit-name {
        font-size: 1.2rem;
        font-weight: 700;
        color: #1a2b6b;
        font-family: 'Barlow Condensed', sans-serif;
        letter-spacing: 0.3px;
        margin-bottom: 0.15rem;
    }
    .unit-sem { font-size: 0.75rem; color: #aaa; margin-bottom: 1rem; }

    div[data-testid="stButton"] button {
        background: #1a2b6b !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        width: 100% !important;
    }
    div[data-testid="stButton"] button:hover { background: #e32119 !important; }

    .sem-badge {
        display: inline-block;
        background: #f0f3ff;
        color: #1a2b6b;
        padding: 0.2rem 0.9rem;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 600;
        margin-bottom: 1.5rem;
    }
    .section-title {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 1.1rem;
        font-weight: 700;
        color: #1a2b6b;
        margin-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Header
def img_to_b64(path):
    try:
        return base64.b64encode(Path(path).read_bytes()).decode()
    except:
        return None

logo_b64  = img_to_b64("logo.png")
logo_html = f'<img src="data:image/png;base64,{logo_b64}" />' if logo_b64 else "🏫"

st.markdown(f"""
<div class="header">
    {logo_html}
    <div class="header-text">
        <h1>Cultura Inglesa — Mapa de Notas</h1>
        <p>Extração automática de notas via Sponte API</p>
    </div>
</div>
""", unsafe_allow_html=True)

# Cards das unidades
cols = st.columns(4)
for idx, (nome, config) in enumerate(UNIDADES.items()):
    with cols[idx]:
        st.markdown(f"""
        <div class="unit-card">
            <div class="unit-name">{nome}</div>
            <div class="unit-sem">Semestre {config["semestre"]}</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("📥 Nova Planilha", key=f"novo_{idx}"):
            st.session_state[f"acao_{idx}"] = "nova"
        if st.button("🔄 Atualizar Existente", key=f"atualizar_{idx}"):
            st.session_state[f"acao_{idx}"] = "atualizar"

# Processamento
for idx, (nome, config) in enumerate(UNIDADES.items()):
    acao = st.session_state.get(f"acao_{idx}")
    if not acao:
        continue

    st.markdown("---")
    st.markdown(f"### {'📥 Nova planilha' if acao == 'nova' else '🔄 Atualizar planilha'} — {nome}")

    arquivo_existente = None
    if acao == "atualizar":
        uploaded = st.file_uploader(
            "Faça upload da planilha anterior para preservar os comentários:",
            type=["xlsx"],
            key=f"upload_{idx}"
        )
        if not uploaded:
            st.info("📎 Aguardando upload da planilha existente...")
            continue
        arquivo_existente = uploaded

    if st.button("▶️ Iniciar geração", key=f"iniciar_{idx}"):
        progress_bar = st.progress(0)
        status_text  = st.empty()

        dados, erro = coletar_dados(config["token"], config["semestre"], progress_bar, status_text)

        if erro:
            st.error(f"❌ {erro}")
        else:
            if arquivo_existente:
                status_text.text("🔄 Mesclando com planilha existente...")
                dados = mesclar_com_existente(dados, arquivo_existente)

            status_text.text("✅ Gerando arquivo Excel...")
            arquivo_final = exportar_excel(dados)

            sem_fmt    = config["semestre"].replace("/", ".")
            nome_arq   = f"mapadenotas_{nome}_{sem_fmt}.xlsx"

            st.success(f"✅ Planilha gerada com sucesso!")
            st.download_button(
                label=f"⬇️ Baixar {nome_arq}",
                data=arquivo_final,
                file_name=nome_arq,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{idx}"
            )
        st.session_state[f"acao_{idx}"] = None
