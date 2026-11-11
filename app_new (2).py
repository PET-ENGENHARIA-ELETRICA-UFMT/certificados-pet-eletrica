"""
Gerador de Certificados — PET Engenharia Elétrica (UFMT)

Interface Streamlit. Toda a lógica (leitura da planilha, regras de casos
específicos, validação, PDF) está em certgen.py.
"""
import base64
import io
import os
from datetime import date
from functools import lru_cache

import pandas as pd
import streamlit as st

import certgen as cg

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO_WHITE = os.path.join(HERE, "Logo 10 Anos - Letra Branca.png")
TEMPLATE_PATH = os.path.join(HERE, "template_padrao.pdf")

MAX_LOGO_PIXELS = 40_000_000   # acima disso o arquivo é pesado demais para decodificar a cada acesso


@lru_cache(maxsize=1)
def load_logo():
    """Logo reduzido (altura 120 px) como (PIL.Image, base64). Lido uma vez por processo.
    Se o arquivo for gigante (ex.: 155 milhões de pixels), não decodifica — evita estourar a memória."""
    if not os.path.exists(LOGO_WHITE):
        return None, ""
    try:
        from PIL import Image

        with Image.open(LOGO_WHITE) as im:          # abre só o cabeçalho; ainda não decodifica
            if im.width * im.height > MAX_LOGO_PIXELS:
                return None, ""
            im = im.convert("RGBA")
        im.thumbnail((600, 120))
        buf = io.BytesIO()
        im.save(buf, format="PNG", optimize=True)
        return im, base64.b64encode(buf.getvalue()).decode()
    except Exception:
        return None, ""


_logo_img, _logo_b64 = load_logo()

st.set_page_config(
    page_title="Gerador de Certificados PET",
    page_icon=_logo_img if _logo_img is not None else "🎓",
    layout="wide",
)

# --------------------------------------------------------------------------
# Estilo
# --------------------------------------------------------------------------
st.markdown(
    """
    <style>
      .block-container {padding-top: 1.6rem; max-width: 1180px;}
      .hero {background: linear-gradient(120deg,#0b1a38 0%,#1b2a4a 60%,#2a3f6e 100%);
             border-radius: 16px; padding: 22px 28px; display:flex; align-items:center; gap:20px;
             border-bottom: 4px solid #c8782a; margin-bottom: 1.1rem;}
      .hero img {height: 58px;}
      .hero h1 {color:#fff; font-size: 1.65rem; margin:0; padding:0; line-height:1.2;}
      .hero p {color:#c9d3ea; margin:.2rem 0 0 0; font-size:.95rem;}
      div[data-testid="stMetric"] {background:#f5f7fb; border:1px solid #e3e8f2; border-radius:12px; padding:10px 14px;}
      div[data-testid="stMetric"] * {color:#1b2a4a;}
      .stTabs [data-baseweb="tab"] {font-weight:600;}
      .hint {color:#5b6478; font-size:.88rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


def logo_b64() -> str:
    return _logo_b64


def hero():
    b64 = logo_b64()
    logo = f'<img src="data:image/png;base64,{b64}"/>' if b64 else ""
    st.markdown(
        f"""<div class="hero">{logo}<div><h1>Gerador de Certificados · PET Elétrica</h1>
        <p>Carregue a planilha (ou preencha direto aqui), confira e baixe todos os PDFs em um ZIP.</p></div></div>""",
        unsafe_allow_html=True,
    )


# --------------------------------------------------------------------------
# Estado
# --------------------------------------------------------------------------
def init_state():
    ss = st.session_state
    defaults = {
        "src_id": None,
        "cursos_df": cg.empty_cursos_df(),
        "people_df": cg.empty_people_df(),
        "people_ver": 0,
        "prog_ver": 0,
        "load_notes": [],
        "course_idx": None,
        "c_nome": "",
        "c_carga": "",
        "c_ini": date.today(),
        "c_fim": date.today(),
        "c_min": "",
        "prog_df": cg.program_to_df([]),
        "preview": None,
        "gen": None,
        "s_frase": cg.DEFAULT_PHRASE,
        "s_vp": cg.DEFAULT_VERBS["participante"],
        "s_vm": cg.DEFAULT_VERBS["ministrante"],
    }
    for k, v in defaults.items():
        ss.setdefault(k, v)


def load_course(idx):
    """Carrega a linha idx da aba 'cursos' nos campos editáveis."""
    ss = st.session_state
    row = ss.cursos_df.iloc[idx]
    ss.course_idx = idx
    ss.c_nome = cg._text(row.get("CURSO"))
    ss.c_carga = cg.format_hours(cg.parse_hours(row.get("CARGAHORARIA")))
    ss.c_ini = cg.parse_date(row.get("DATAINICIO")) or date.today()
    ss.c_fim = cg.parse_date(row.get("DATAFIM")) or date.today()
    ss.c_min = "\n".join(cg.split_list(row.get("MINISTRANTES")))
    ss.prog_df = cg.program_to_df(cg.parse_program(row.get("PROGRAMA")))
    ss.prog_ver += 1


def on_course_change():
    load_course(st.session_state.course_sel)


def new_course():
    ss = st.session_state
    ss.course_idx = None
    ss.c_nome, ss.c_carga, ss.c_min = "", "", ""
    ss.c_ini = ss.c_fim = date.today()
    ss.prog_df = cg.program_to_df([])
    ss.prog_ver += 1


def swap_dates():
    ss = st.session_state
    sw = cg.swapped_dates(ss.c_ini, ss.c_fim)
    if sw:
        ss.c_ini, ss.c_fim = sw


def import_program_text():
    ss = st.session_state
    ss.prog_df = cg.program_to_df(cg.parse_program(ss.get("prog_paste", "")))
    ss.prog_ver += 1
    ss.prog_paste = ""


def reset_texts():
    ss = st.session_state
    ss.s_frase = cg.DEFAULT_PHRASE
    ss.s_vp = cg.DEFAULT_VERBS["participante"]
    ss.s_vm = cg.DEFAULT_VERBS["ministrante"]


def course_options() -> list:
    """Nomes de curso para a coluna CURSO da lista de pessoas (planilha + curso aberto + valores já usados)."""
    ss = st.session_state
    names = list(ss.cursos_df["CURSO"]) if len(ss.cursos_df) else []
    names.append(ss.get("c_nome", ""))
    names += list(ss.people_df["CURSO"])
    names = (cg._clean_spaces(cg._text(n)) for n in names)
    return list(dict.fromkeys(n for n in names if n))


def pdf_pages(pdf_bytes: bytes):
    """Páginas do PDF como imagens (pré-visualização). None se pypdfium2 não estiver instalado."""
    try:
        import pypdfium2 as pdfium
    except ImportError:
        return None
    doc = pdfium.PdfDocument(pdf_bytes)
    return [doc[i].render(scale=1.6).to_pil() for i in range(len(doc))]


def get_template():
    up = st.session_state.get("tpl_upload")
    if up is not None:
        return up.getvalue(), "personalizado"
    if os.path.exists(TEMPLATE_PATH):
        with open(TEMPLATE_PATH, "rb") as f:
            return f.read(), "padrão"
    return None, None


# --------------------------------------------------------------------------
# Barra lateral
# --------------------------------------------------------------------------
init_state()
ss = st.session_state
with st.sidebar:
    if logo_b64():
        st.markdown(
            f'<div style="background:#1b2a4a;border-radius:12px;padding:10px;text-align:center">'
            f'<img style="height:44px" src="data:image/png;base64,{logo_b64()}"/></div>',
            unsafe_allow_html=True,
        )
    st.subheader("Como usar")
    st.markdown(
        "1. **Curso** — envie a planilha ou preencha aqui\n"
        "2. **Pessoas** — confira a lista; ajuste carga e partes cursadas de quem fez só uma parte\n"
        "3. **Emissão** — data, numeração, textos\n"
        "4. **Revisar e gerar** — veja os alertas, pré-visualize e baixe o ZIP"
    )
    st.download_button(
        "📥 Baixar planilha modelo",
        data=cg.make_template_workbook(),
        file_name="Modelo_certificado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        help="Com instruções, listas suspensas e validação nas células.",
    )
    st.divider()
    st.file_uploader("📄 Template PDF personalizado (opcional)", type=["pdf"], key="tpl_upload")
    tpl_bytes, tpl_kind = get_template()
    if tpl_bytes:
        st.caption(f"Usando template {tpl_kind}.")
    else:
        st.error("Template não encontrado. Envie um PDF acima.")

hero()
tab_curso, tab_pessoas, tab_emissao, tab_final = st.tabs(
    ["1 · Curso", "2 · Pessoas", "3 · Emissão", "4 · Revisar e gerar"]
)

# --------------------------------------------------------------------------
# 1 · Curso
# --------------------------------------------------------------------------
with tab_curso:
    uploaded = st.file_uploader("📂 Planilha Excel (.xlsx) — opcional", type=["xlsx"], key="xlsx_upload")
    if uploaded is not None:
        fid = (uploaded.name, uploaded.size)
        if ss.src_id != fid:
            try:
                cursos_df, people_df, notes = cg.read_workbook(uploaded.getvalue())
                ss.src_id = fid
                ss.cursos_df, ss.people_df, ss.load_notes = cursos_df, people_df, notes
                ss.people_ver += 1
                ss.gen = ss.preview = None
                if len(cursos_df):
                    ss.course_sel = 0
                    load_course(0)
                else:
                    new_course()
            except Exception as e:  # planilha corrompida / formato inesperado
                st.error(f"Não consegui ler a planilha: {e}")
    for n in ss.load_notes:
        {"erro": st.error, "aviso": st.warning}.get(n.level, st.info)(n.msg)

    top1, top2 = st.columns([3, 1], vertical_alignment="bottom")
    if len(ss.cursos_df):
        top1.selectbox(
            "Curso da planilha",
            options=list(range(len(ss.cursos_df))),
            format_func=lambda i: cg._text(ss.cursos_df.iloc[i]["CURSO"]),
            key="course_sel",
            on_change=on_course_change,
        )
        top2.button("➕ Novo curso em branco", on_click=new_course, width="stretch")
    else:
        top1.caption("Sem planilha: preencha os dados do curso abaixo (ou envie a planilha acima).")

    st.text_input("Nome do curso", key="c_nome", placeholder="Ex.: Minicurso de LaTeX")
    c1, c2, c3 = st.columns(3)
    c1.text_input("Carga horária (horas)", key="c_carga", placeholder="30", help="Aceita 30, 2,5 ou 2h30.")
    c2.date_input("Data de início", key="c_ini", format="DD/MM/YYYY")
    c3.date_input("Data de fim", key="c_fim", format="DD/MM/YYYY")
    if ss.c_carga.strip() and cg.parse_hours(ss.c_carga) is None:
        st.error(f"Não entendi a carga horária “{ss.c_carga}”. Use um número, como 30 ou 2,5.")
    if ss.c_fim < ss.c_ini:
        st.error("A data de fim é anterior à data de início.")
    elif cg.swapped_dates(ss.c_ini, ss.c_fim):
        a, b = cg.swapped_dates(ss.c_ini, ss.c_fim)
        w1, w2 = st.columns([4, 1], vertical_alignment="center")
        w1.warning(
            f"O período ({(ss.c_fim - ss.c_ini).days + 1} dias) parece longo — dia e mês podem estar trocados. "
            f"Seria {a:%d/%m/%Y} a {b:%d/%m/%Y}?"
        )
        w2.button("🔁 Trocar dia/mês", on_click=swap_dates, width="stretch")

    st.markdown("##### Conteúdo programático")
    st.markdown(
        '<p class="hint">Cada linha é um item da página 2. <b>Título</b> aparece em negrito (ex.: “1º Dia”); '
        "<b>Tópico</b> aparece com marcador. Use a linha em branco no fim para adicionar; selecione uma linha para apagar.</p>",
        unsafe_allow_html=True,
    )
    prog_edit = st.data_editor(
        ss.prog_df,
        key=f"prog_ed_{ss.prog_ver}",
        num_rows="dynamic",
        hide_index=True,
        width="stretch",
        column_config={
            "TIPO": st.column_config.SelectboxColumn("Tipo", options=["Tópico", "Título"], default="Tópico", required=True, width="small"),
            "TEXTO": st.column_config.TextColumn("Texto", width="large"),
        },
    )
    with st.expander("📋 Colar conteúdo em texto (itens separados por ;)"):
        st.text_area(
            "Texto",
            key="prog_paste",
            height=110,
            placeholder="1º Dia: Fundamentos;- Introdução ao LaTeX;- Equações;2º Dia: Gráficos;- Imagens",
            help="Itens iniciados por “-” ou “•” viram tópicos. Linhas que terminam em “:” ou começam com “1º Dia”, “Módulo 2”… viram títulos.",
        )
        st.button("Importar para a tabela (substitui o atual)", on_click=import_program_text)

    st.text_area(
        "Ministrantes / organizadores",
        key="c_min",
        height=130,
        help="Um nome por linha (ou separados por ;). Aparecem na página 2 e recebem certificado de ministrante.",
    )

# --------------------------------------------------------------------------
# 2 · Pessoas
# --------------------------------------------------------------------------
prog_items = cg.df_to_program(prog_edit)

with tab_pessoas:
    st.markdown(
        '<p class="hint">Deixe <b>Carga horária</b>, <b>Partes cursadas</b> e <b>Nº</b> em branco para usar o padrão do curso / numeração automática. '
        "Ministrantes cadastrados no curso entram sozinhos — só liste aqui para um ajuste individual. "
        "A coluna <b>Curso</b> separa as pessoas quando a planilha tem vários cursos (em branco = vale para o curso aberto).</p>",
        unsafe_allow_html=True,
    )
    sections_now = cg.program_sections(prog_items)
    if sections_now:
        with st.expander(f"📑 Partes do curso (para a coluna “Partes cursadas”) — {len(sections_now)}", expanded=False):
            st.markdown("\n".join(f"**{i}.** {label}" for i, (label, _) in enumerate(sections_now, start=1)))
            st.caption("Na coluna “Partes cursadas” digite, por exemplo: 1  ·  1;3  ·  1-2  ·  ou parte do título (1º Dia).")
    else:
        st.caption("Cadastre o conteúdo do curso na aba 1 para poder indicar partes cursadas.")
    with st.expander("➕ Colar vários nomes de uma vez"):
        qa1, qa2 = st.columns([3, 1])
        qa1.text_area("Nomes (um por linha)", key="qa_text", height=120, placeholder="Cole aqui a coluna de nomes do Excel/Forms")
        qa2.selectbox("Tipo", ["Participante", "Ministrante"], key="qa_tipo")
        qa_btn = qa2.button("Adicionar à lista", width="stretch")

    people_edit = st.data_editor(
        ss.people_df,
        key=f"people_ed_{ss.people_ver}",
        num_rows="dynamic",
        hide_index=True,
        width="stretch",
        column_order=cg.PEOPLE_COLS,
        column_config={
            "NOME": st.column_config.TextColumn("Nome completo", required=True, width="large"),
            "TIPO": st.column_config.SelectboxColumn("Tipo", options=["Participante", "Ministrante"], default="Participante", width="small"),
            "CURSO": st.column_config.SelectboxColumn("Curso", options=course_options(), help="Vazio = curso aberto no site. Útil quando a planilha tem vários cursos.", width="medium"),
            "CARGAHORARIA": st.column_config.NumberColumn("Carga horária (h)", min_value=0.0, step=0.5, format="%g", help="Vazio = carga horária do curso.", width="small"),
            "PARTES": st.column_config.TextColumn("Partes cursadas", help="Vazio = conteúdo completo. Ex.: 1  ·  1;3  ·  1-2  ·  1º Dia. Veja a numeração em “Partes do curso”.", width="small"),
            "CERT_NUM": st.column_config.TextColumn("Nº do certificado", help="Vazio = numeração automática.", width="small"),
            "PROGRAMA": st.column_config.TextColumn("Conteúdo por texto", help="Opcional. Alternativa a “Partes cursadas”: o conteúdo em texto (itens separados por ;).", width="medium"),
        },
    )
    _ck = cg._norm_key(ss.c_nome)
    if _ck and len(people_edit):
        _outros = sum(1 for v in people_edit["CURSO"] if cg._norm_key(v) and cg._norm_key(v) != _ck)
        st.caption(f"{len(people_edit) - _outros} linha(s) valem para “{ss.c_nome.strip()}”" + (f" · {_outros} são de outros cursos e ficam de fora." if _outros else "."))
    _sem = cg.people_without_course(people_edit, ss.cursos_df)
    if _sem:
        st.warning("Curso não encontrado na planilha (erro de digitação?): " + cg._fmt_names(_sem, 5))

    def _commit(df):
        ss.people_df = df.reset_index(drop=True)
        ss.people_ver += 1
        st.rerun()

    if qa_btn:
        nomes = [n.strip() for n in ss.qa_text.splitlines() if n.strip()]
        if nomes:
            curso_padrao = ss.c_nome.strip() if len(ss.cursos_df) > 1 else ""   # com vários cursos, já marca o aberto
            novos = pd.DataFrame(
                {"NOME": nomes, "TIPO": ss.qa_tipo, "CURSO": curso_padrao or None, "CARGAHORARIA": float("nan"),
                 "PARTES": "", "CERT_NUM": "", "PROGRAMA": ""}
            )
            ss.qa_text = ""
            _commit(pd.concat([people_edit, novos], ignore_index=True))

    with st.expander("⚡ Assistente de participação parcial (carga horária + partes cursadas, em lote)"):
        base = people_edit.reset_index(drop=True)
        sections = cg.program_sections(prog_items)
        pick = st.multiselect(
            "Pessoas",
            options=list(range(len(base))),
            format_func=lambda i: f"{i + 1}. {cg._text(base.iloc[i]['NOME']) or '(sem nome)'}",
            help="Selecione quem fez a mesma coisa (ex.: todos os que ficaram só 8 h) e aplique de uma vez.",
        )
        a1, a2 = st.columns([1, 2])
        horas_txt = a1.text_input("Carga horária", placeholder="vazio = não alterar", help="Aceita 8, 2,5 ou 2h30.")
        cursado = a2.multiselect(
            "Partes cursadas",
            options=list(range(len(sections))),
            format_func=lambda i: f"{i + 1}. {sections[i][0]}",
            placeholder="vazio = não alterar" if sections else "cadastre o conteúdo do curso na aba 1",
        )
        horas = cg.parse_hours(horas_txt) if horas_txt.strip() else None
        if horas_txt.strip() and horas is None:
            st.error(f"Não entendi a carga horária “{horas_txt}”. Use um número, como 8 ou 2,5.")
        b1, b2, _ = st.columns([1.3, 1.3, 2])
        if b1.button(
            "Aplicar às selecionadas", type="primary", width="stretch",
            disabled=not pick or (horas is None and not cursado) or (bool(horas_txt.strip()) and horas is None),
        ):
            df = base.copy()
            partes = ";".join(str(i + 1) for i in sorted(cursado))
            for i in pick:
                if horas is not None:
                    df.at[i, "CARGAHORARIA"] = float(horas)
                if cursado:
                    df.at[i, "PARTES"] = partes
                    df.at[i, "PROGRAMA"] = ""      # PARTES substitui o texto antigo
            _commit(df)
        if b2.button("Limpar personalização", disabled=not pick, width="stretch", help="Volta carga horária e conteúdo ao padrão do curso."):
            df = base.copy()
            for i in pick:
                df.at[i, "CARGAHORARIA"] = float("nan")
                df.at[i, "PARTES"] = ""
                df.at[i, "PROGRAMA"] = ""
            _commit(df)

# --------------------------------------------------------------------------
# 3 · Emissão
# --------------------------------------------------------------------------
with tab_emissao:
    e1, e2, e3 = st.columns(3)
    alvo = e1.radio("Gerar para", ["Participantes", "Ministrantes", "Ambos"], key="s_alvo", horizontal=True)
    data_ass = e2.date_input("Data da assinatura", value=date.today(), key="s_data", format="DD/MM/YYYY")
    local = e3.text_input("Local", value="Cuiabá-MT", key="s_local")
    e4, e5, e6 = st.columns(3)
    num_ini = e4.number_input("Numeração inicial", min_value=1, value=1, step=1, key="s_num")
    volume = e5.text_input("Volume do livro (opcional)", key="s_vol")
    padr = e6.checkbox(
        "Padronizar nomes em MAIÚSCULAS/minúsculas", value=True, key="s_padr",
        help="“MARIA DA SILVA” vira “Maria da Silva”. Nomes já com caixa mista não mudam.",
    )
    with st.expander("✏️ Textos do certificado (frase da página 1)"):
        st.text_area(
            "Frase",
            key="s_frase",
            height=90,
            help="Campos: {nome} {papel} {curso} {periodo} {carga}. Aceita <b>negrito</b>.",
        )
        v1, v2 = st.columns(2)
        v1.text_input("Verbo para participantes ({papel})", key="s_vp")
        v2.text_input("Verbo para ministrantes ({papel})", key="s_vm")
        st.button("Restaurar textos padrão", on_click=reset_texts)

# --------------------------------------------------------------------------
# 4 · Revisar e gerar
# --------------------------------------------------------------------------
with tab_final:
    course = cg.Course(
        nome=ss.c_nome.strip(),
        carga=cg.parse_hours(ss.c_carga),
        dt_ini=ss.c_ini,
        dt_fim=ss.c_fim,
        programa=prog_items,
        ministrantes=cg.split_list(ss.c_min),
    )
    settings = cg.Settings(
        local=local.strip() or "Cuiabá-MT",
        data_assinatura=data_ass,
        volume=volume.strip(),
        num_inicial=int(num_ini),
        frase=ss.s_frase,
        verbos={"participante": ss.s_vp, "ministrante": ss.s_vm},
        padronizar_nomes=padr,
    )
    targets = {
        "Participantes": {"participante"},
        "Ministrantes": {"ministrante"},
        "Ambos": {"participante", "ministrante"},
    }[alvo]
    recs, rec_issues = cg.build_recipients(course, people_edit, settings, targets)
    issues = cg.validate(course, recs, settings, rec_issues, template_ok=tpl_bytes is not None)
    erros = [i for i in issues if i.level == "erro"]
    avisos = [i for i in issues if i.level == "aviso"]
    infos = [i for i in issues if i.level == "info"]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Certificados", len(recs))
    m2.metric("Participantes", sum(r.role == "participante" for r in recs))
    m3.metric("Ministrantes", sum(r.role == "ministrante" for r in recs))
    m4.metric("Com ajuste individual", sum(r.parcial for r in recs))

    if erros:
        st.error("**Corrija antes de gerar:**\n\n" + "\n".join(f"- {i.msg}" for i in erros))
    if avisos:
        st.warning("**Confira:**\n\n" + "\n".join(f"- {i.msg}" for i in avisos))
    if infos:
        st.info("\n".join(f"- {i.msg}" for i in infos))
    if not issues:
        st.success("Tudo certo — nenhum alerta.")

    if recs:
        with st.expander(f"👥 Quem vai receber ({len(recs)})"):
            tabela = pd.DataFrame(
                {
                    "Nº": [r.numero for r in recs],
                    "Nome": [r.nome for r in recs],
                    "Tipo": [cg.ROLE_LABEL[r.role] for r in recs],
                    "Carga horária": [cg.hours_phrase(r.carga if r.carga is not None else course.carga) for r in recs],
                    "Conteúdo": [
                        (f"Partes {cg.parts_label(r.partes)}" if r.partes else "Parcial") if r.programa else "Completo"
                        for r in recs
                    ],
                }
            )
            st.dataframe(tabela, hide_index=True, width="stretch")

    st.divider()
    p1, p2 = st.columns([2, 1], vertical_alignment="bottom")
    prev_idx = None
    if recs:
        prev_idx = p1.selectbox(
            "Pré-visualizar certificado de",
            options=list(range(len(recs))),
            format_func=lambda i: f"{recs[i].numero} · {recs[i].nome}",
        )
    if p2.button("👁️ Pré-visualizar", disabled=not recs or tpl_bytes is None, width="stretch"):
        try:
            pdf, _ = cg.render_certificate(tpl_bytes, course, recs[prev_idx], settings)
            ss.preview = (recs[prev_idx].nome, pdf)
        except Exception as e:
            st.error(f"Erro na pré-visualização: {e}")
    if ss.preview:
        nome_prev, pdf_prev = ss.preview
        st.caption(f"Pré-visualização: {nome_prev}")
        imgs = pdf_pages(pdf_prev)
        if imgs:
            for img in imgs:
                st.image(img, width="stretch")
        else:
            b64 = base64.b64encode(pdf_prev).decode()
            st.markdown(f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="520"></iframe>', unsafe_allow_html=True)
        st.download_button("⬇️ Baixar este PDF", pdf_prev, file_name=f"{nome_prev}.pdf", mime="application/pdf")

    st.divider()
    g1, g2 = st.columns([2, 1])
    if g1.button(
        "🚀 Gerar todos os certificados (ZIP)", type="primary",
        disabled=bool(erros) or not recs or tpl_bytes is None, width="stretch",
    ):
        bar = st.progress(0.0, text="Gerando certificados…")
        try:
            zbytes, notes = cg.build_zip(
                tpl_bytes, course, recs, settings,
                progress=lambda f: bar.progress(f, text=f"Gerando certificados… {int(f * 100)}%"),
            )
            ss.gen = {"zip": zbytes, "notes": notes, "n": len(recs), "stamp": pd.Timestamp.now().strftime("%Y%m%d_%H%M")}
        except Exception as e:
            ss.gen = None
            st.error(f"Erro ao gerar: {e}")
        bar.empty()
    if ss.gen:
        st.success(f"{ss.gen['n']} certificados gerados. O ZIP inclui também `_registro_certificados.csv` (nº, nome, carga horária, partes).")
        for n in ss.gen["notes"]:
            st.warning(n)
        st.download_button(
            "📦 Baixar ZIP completo",
            ss.gen["zip"],
            file_name=f"Certificados_{ss.gen['stamp']}.zip",
            mime="application/zip",
            type="primary",
        )

    # planilha atualizada (guarda as edições feitas no site, no mesmo formato aceito pelo gerador)
    cursos_out = ss.cursos_df.copy()
    row = {
        "CURSO": ss.c_nome.strip(),
        "DATAINICIO": ss.c_ini,
        "DATAFIM": ss.c_fim,
        "CARGAHORARIA": cg.parse_hours(ss.c_carga),
        "PROGRAMA": cg.program_to_text(prog_items),
        "MINISTRANTES": ";".join(cg.split_list(ss.c_min)),
    }
    if row["CURSO"]:
        idx = ss.course_idx
        if idx is not None and idx < len(cursos_out):
            for k, v in row.items():
                cursos_out.at[idx, k] = v
        else:
            cursos_out = pd.concat([cursos_out, pd.DataFrame([row])], ignore_index=True)
    g2.download_button(
        "💾 Baixar planilha atualizada",
        cg.export_workbook(cursos_out, people_edit),
        file_name="Certificados_planilha_atualizada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        help="Guarda curso, conteúdo e lista de pessoas exatamente como estão agora.",
    )
