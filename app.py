"""
Sponte Notas — App Streamlit
Gera planilha de notas por unidade com um clique.
"""

import streamlit as st
import requests
import pandas as pd
import time
import io
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ─── Configuração das unidades ─────────────────────────────────────────────────
UNIDADES = {
    "🏫 Young":      {"token": "LcC5k0eOQEdY", "codigo": "71976"},
    "🏫 Boa Viagem": {"token": "D2vEp1qCSPpv", "codigo": "71961"},
    "🏫 Setúbal":    {"token": "UO7c764YXSRU", "codigo": "71977"},
    "🏫 Natal":      {"token": "1fvBagK0xYnp", "codigo": "71978"},
}

SEMESTRE  = "2026.1"
SITUATION = 1
BASE_URL  = "https://webservices.sponteweb.com.br/WSApiSponteRest/api"

# ─── Componentes e estrutura ───────────────────────────────────────────────────
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
COLUNAS_EXTRA = ["Action Plan", "Coordinator's Comment"]

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

COLUNAS_A    = build_colunas_a()
COLUNAS_B    = [f"{pre} - {comp}" for _, (pre, comp) in PROVAS_B_SIMPLES.items()] + ["Média Geral Av"]
COLUNAS_INFO_A = ["Turma", "Phase", "Professor", "student_id", "Nome Aluno", "Situação", "Média Geral"]
COLUNAS_INFO_B = ["Turma", "Phase", "Professor", "student_id", "Nome Aluno", "Situação"]
TODAS_COLS_A = COLUNAS_INFO_A + COLUNAS_A + COLUNAS_EXTRA
TODAS_COLS_B = COLUNAS_INFO_B + COLUNAS_B + COLUNAS_EXTRA

# ─── Funções API ───────────────────────────────────────────────────────────────
def headers(token):
    return {"Accept": "application/json", "Content-Type": "application/json", "api_key": token}

def get_phases_map(token):
    r = requests.get(f"{BASE_URL}/phases", headers=headers(token))
    return {p["name"]: p["phase_id"] for p in r.json()} if r.ok else {}

def get_turmas(token):
    r = requests.get(f"{BASE_URL}/classes", headers=headers(token))
    if not r.ok:
        return []
    return [t for t in r.json() if t.get("situation") == SITUATION and SEMESTRE in t.get("name", "")]

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

# ─── Formatação Excel ──────────────────────────────────────────────────────────
FILL_HEADER = PatternFill("solid", fgColor="6200A8")
FILL_RED    = PatternFill("solid", fgColor="FFCCCC")
FILL_YELLOW = PatternFill("solid", fgColor="FFFACD")
FONT_WHITE  = Font(bold=True, color="FFFFFF")

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
    nota_cols = {i+1 for i, c in enumerate(colunas) if is_nota_col(c)}
    for row in ws.iter_rows(min_row=2):
        for cell in row:
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
def gerar_planilha(nome_unidade, token, progress_bar, status_text):
    phases_map = get_phases_map(token)
    turmas     = get_turmas(token)

    if not turmas:
        return None, "Nenhuma turma aberta encontrada."

    dados       = {}
    cache_nomes = {}
    total       = len(turmas)

    for i, turma in enumerate(turmas):
        class_id   = turma["class_id"]
        nome_turma = turma["name"]
        status_text.text(f"⏳ Processando turma {i+1}/{total}: {nome_turma}")
        progress_bar.progress((i) / total)

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
            }
            linha.update(notas_dict)
            for col in todas_cols:
                linha.setdefault(col, "")
            dados[professor][fmt].append(linha)

    progress_bar.progress(1.0)
    status_text.text("✅ Concluído! Gerando arquivo...")

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
    return output, None

# ─── Interface Streamlit ───────────────────────────────────────────────────────
st.set_page_config(
    page_title="Notas — Cultura Inglesa",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&family=DM+Serif+Display&display=swap');

    html, body, [class*="css"] {
        font-family: 'DM Sans', sans-serif;
    }
    .main { background-color: #f8f7ff; }

    .header {
        background: linear-gradient(135deg, #6200a8 0%, #9c27b0 100%);
        padding: 2rem 2.5rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        color: white;
    }
    .header h1 {
        font-family: 'DM Serif Display', serif;
        font-size: 2rem;
        margin: 0;
        color: white;
    }
    .header p {
        margin: 0.25rem 0 0;
        opacity: 0.85;
        font-size: 0.95rem;
    }

    .unit-card {
        background: white;
        border: 1px solid #e8e0f0;
        border-radius: 12px;
        padding: 1.5rem;
        text-align: center;
        transition: box-shadow 0.2s;
    }
    .unit-card:hover { box-shadow: 0 4px 20px rgba(98,0,168,0.12); }
    .unit-name {
        font-size: 1.1rem;
        font-weight: 600;
        color: #2d1b4e;
        margin-bottom: 0.25rem;
    }
    .unit-sub {
        font-size: 0.8rem;
        color: #888;
        margin-bottom: 1rem;
    }
    .stButton > button {
        background: #6200a8 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
        width: 100% !important;
    }
    .stButton > button:hover {
        background: #7b00d4 !important;
    }
    .semestre-badge {
        display: inline-block;
        background: #f0e6ff;
        color: #6200a8;
        padding: 0.2rem 0.8rem;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 600;
        margin-bottom: 1.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown(f"""
<div class="header">
    <h1>📊 Cultura Inglesa — Notas</h1>
    <p>Extração automática de notas via Sponte API</p>
</div>
""", unsafe_allow_html=True)

st.markdown(f'<div class="semestre-badge">Semestre {SEMESTRE}</div>', unsafe_allow_html=True)

# Cards das unidades
cols = st.columns(4)
for idx, (nome, config) in enumerate(UNIDADES.items()):
    with cols[idx]:
        nome_limpo = nome.replace("🏫 ", "")
        st.markdown(f"""
        <div class="unit-card">
            <div class="unit-name">{nome_limpo}</div>
            <div class="unit-sub">Cód. {config['codigo']}</div>
        </div>
        """, unsafe_allow_html=True)

        if st.button(f"Gerar Planilha", key=f"btn_{idx}"):
            st.session_state[f"gerar_{idx}"] = True

# Processamento
for idx, (nome, config) in enumerate(UNIDADES.items()):
    if st.session_state.get(f"gerar_{idx}"):
        st.session_state[f"gerar_{idx}"] = False
        nome_limpo = nome.replace("🏫 ", "")

        st.markdown("---")
        st.markdown(f"### Gerando planilha — {nome_limpo}")

        progress_bar = st.progress(0)
        status_text  = st.empty()

        arquivo, erro = gerar_planilha(nome_limpo, config["token"], progress_bar, status_text)

        if erro:
            st.error(f"❌ {erro}")
        else:
            st.success(f"✅ Planilha de {nome_limpo} gerada com sucesso!")
            st.download_button(
                label=f"⬇️ Baixar notas_{nome_limpo}_{SEMESTRE}.xlsx",
                data=arquivo,
                file_name=f"notas_{nome_limpo}_{SEMESTRE}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{idx}"
            )
