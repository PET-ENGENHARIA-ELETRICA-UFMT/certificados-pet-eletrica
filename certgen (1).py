"""
certgen.py — lógica do Gerador de Certificados do PET Engenharia Elétrica (UFMT).

A interface (Streamlit) fica em app.py; aqui estão:
  • leitura tolerante da planilha (datas em vários formatos, carga "2h30", ementa livre,
    aba "conteudo" com um item por linha, coluna CURSO e coluna PARTES nas pessoas…)
  • regras dos casos específicos (carga/conteúdo parcial, nº manual, ministrantes extras…)
  • validação com alertas em português
  • geração dos PDFs (reportlab + pypdf) com ajuste automático de fonte
  • planilha modelo e exportação da planilha atualizada
"""
from __future__ import annotations

import csv
import html
import io
import math
import re
import unicodedata
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from functools import lru_cache

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.platypus import Frame, Paragraph

try:  # pypdf é o sucessor do PyPDF2
    from pypdf import PdfReader, PdfWriter
except ImportError:  # pragma: no cover
    from PyPDF2 import PdfReader, PdfWriter

# ==========================================================================
# Constantes
# ==========================================================================
CURSOS_COLS = ["CURSO", "DATAINICIO", "DATAFIM", "CARGAHORARIA", "PROGRAMA", "MINISTRANTES"]
PEOPLE_COLS = ["NOME", "TIPO", "CURSO", "CARGAHORARIA", "PARTES", "CERT_NUM", "PROGRAMA"]
CONTENT_COLS = ["CURSO", "TIPO", "TEXTO"]
ROLE_LABEL = {"participante": "Participante", "ministrante": "Ministrante"}
KIND_LABEL = {"titulo": "Título", "topico": "Tópico"}
EXAMPLE_PREFIX = "EXEMPLO"          # linhas de exemplo da planilha modelo são ignoradas na leitura
FAIL_PREFIX = "Não gerado: "        # prefixo das notas de build_zip para certificados que falharam

DEFAULT_PHRASE = (
    "Certificamos que <b>{nome}</b> {papel} <b>{curso}</b>, {periodo}, "
    "com carga horária de {carga}."
)
DEFAULT_VERBS = {"participante": "participou do", "ministrante": "organizou o"}

MAX_NORMAL_SPAN_DAYS = 31           # acima disso o período "parece longo" (possível dia/mês trocado)

_PARTICLES = {"da", "de", "do", "das", "dos", "e", "di", "du", "del", "la", "le", "van", "von", "der"}
_ROMAN = {"ii", "iii", "iv", "vi", "vii", "viii", "ix"}
_BULLETS = "-–—•·*"
_TITLE_RE = re.compile(
    r"^(?:\d+\s*[ºo°ª]?\s*(?:dia|módulo|modulo|etapa|parte|semana|aula|encontro|turno|sessão|sessao)\b"
    r"|(?:dia|módulo|modulo|etapa|parte|semana|aula|encontro|turno|sessão|sessao)\s*\d+)",
    re.IGNORECASE,
)


# ==========================================================================
# Tipos
# ==========================================================================
@dataclass
class Note:
    level: str   # "erro" | "aviso" | "info"
    msg: str


Issue = Note


@dataclass(frozen=True)
class ProgItem:
    kind: str    # "titulo" | "topico"
    text: str


@dataclass
class Course:
    nome: str = ""
    carga: float | None = None
    dt_ini: date | None = None
    dt_fim: date | None = None
    programa: list = field(default_factory=list)       # list[ProgItem]
    ministrantes: list = field(default_factory=list)   # list[str]


@dataclass
class Settings:
    local: str = "Cuiabá-MT"
    data_assinatura: date = field(default_factory=date.today)
    volume: str = ""
    num_inicial: int = 1
    frase: str = DEFAULT_PHRASE
    verbos: dict = field(default_factory=lambda: dict(DEFAULT_VERBS))
    padronizar_nomes: bool = True


@dataclass
class Recipient:
    nome: str
    role: str                          # "participante" | "ministrante"
    numero: str = ""
    carga: float | None = None         # None = usa a do curso
    programa: list | None = None       # None/[] = conteúdo completo do curso
    parcial: bool = False              # tem ajuste individual de carga e/ou conteúdo
    partes: list = field(default_factory=list)   # nº das partes do curso cursadas (se veio de PARTES)


# ==========================================================================
# Texto, números e datas
# ==========================================================================
def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str):
        return not v.strip()
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def _text(v) -> str:
    """Valor de célula → texto limpo. None/NaN → ''. 12.0 → '12'."""
    if _is_blank(v):
        return ""
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    return str(v).replace("\u00a0", " ").strip()


def _clean_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s.replace("\u00a0", " ")).strip()


def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def _norm_key(s) -> str:
    """Chave de comparação: sem acento, minúscula, espaços únicos."""
    return _clean_spaces(_strip_accents(_text(s))).casefold()


def _esc(s) -> str:
    return html.escape(str(s), quote=False)


def _cap_word(w: str) -> str:
    if w in _ROMAN:
        return w.upper()
    return re.sub(r"(^|[-'’])(\w)", lambda m: m.group(1) + m.group(2).upper(), w)


def smart_title(name: str) -> str:
    """'MARIA DA SILVA' → 'Maria da Silva' (partículas ficam minúsculas)."""
    out = []
    for i, w in enumerate(name.split(" ")):
        lw = w.lower()
        out.append(lw if (i > 0 and lw in _PARTICLES) else _cap_word(lw))
    return " ".join(out)


def clean_name(name, padronizar: bool = True) -> str:
    n = _clean_spaces(_text(name)).strip(";,")
    if n.endswith(".") and len(n.rstrip(".").split(" ")[-1]) > 3:
        n = n.rstrip(".")
    if padronizar and n and (n.isupper() or n.islower()):
        n = smart_title(n)
    return n


def parse_hours(v) -> float | None:
    """Aceita 30, '2,5', '2h30', '2h', '1:30', '90 min', '30 horas'. Vazio/ilegível/≤0 → None."""
    if _is_blank(v) or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        x = float(v)
        return x if math.isfinite(x) and x > 0 else None
    s = _text(v).lower()
    m = re.fullmatch(r"(\d+)\s*h(?:oras?|rs?|s)?\s*(\d{1,2})?\s*(?:min(?:utos?)?|m)?", s)
    if m:
        x = float(int(m.group(1))) + (int(m.group(2)) / 60 if m.group(2) else 0)
        return round(x, 4) if x > 0 else None
    m = re.fullmatch(r"(\d+)\s*:\s*(\d{1,2})", s)
    if m:
        x = float(int(m.group(1))) + int(m.group(2)) / 60
        return round(x, 4) if x > 0 else None
    m = re.fullmatch(r"([\d.,]+)\s*(?:min(?:utos?)?)", s)
    unit_min = bool(m)
    if not m:
        m = re.fullmatch(r"([\d.,]+)\s*(?:h|hs|hr|hrs|horas?)?\.?", s)
    if not m:
        return None
    num = m.group(1)
    if "," in num:
        num = num.replace(".", "").replace(",", ".")
    try:
        x = float(num)
    except ValueError:
        return None
    if unit_min:
        x /= 60
    return round(x, 4) if (math.isfinite(x) and x > 0) else None


def format_hours(x) -> str:
    """30.0 → '30'; 2.5 → '2,5'; None → ''."""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    x = float(x)
    if x.is_integer():
        return str(int(x))
    return f"{x:.2f}".rstrip("0").rstrip(".").replace(".", ",")


def hours_phrase(x) -> str:
    """30 → '30 horas'; 1 → '1 hora'; 2.5 → '2 horas e 30 minutos'; None → '—'."""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return "—"
    x = float(x)
    if x.is_integer():
        n = int(x)
        return f"{n} hora" if n == 1 else f"{n} horas"
    h, frac = int(x), x - int(x)
    mins = round(frac * 60)
    if abs(mins - frac * 60) < 0.01 and mins < 60:
        if h == 0:
            return f"{mins} minutos"
        return f"{h} hora{'s' if h != 1 else ''} e {mins} minutos"
    return f"{format_hours(x)} horas"


def parse_date(v) -> date | None:
    """Aceita date/datetime/Timestamp, nº serial do Excel e textos dd/mm/aaaa, dd.mm.aa, aaaa-mm-dd."""
    if _is_blank(v):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 20000 <= float(v) <= 80000:
            return (datetime(1899, 12, 30) + timedelta(days=float(v))).date()
        return None
    s = _text(v)
    m = re.match(r"^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2}|\d{4})(?:[ T].*)?$", s)
    try:
        if m:
            d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            return date(y + 2000 if y < 100 else y, mo, d)
        m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?$", s)
        if m:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None
    return None


def _swap_dm(d: date) -> date | None:
    if d.day > 12:
        return None
    try:
        return date(d.year, d.day, d.month)
    except ValueError:
        return None


def swapped_dates(ini, fim):
    """Se o período é negativo ou muito longo e trocar dia/mês o torna razoável, devolve (ini, fim) sugeridos."""
    if not ini or not fim:
        return None
    span = (fim - ini).days
    if 0 <= span <= MAX_NORMAL_SPAN_DAYS:
        return None
    si, sf = _swap_dm(ini), _swap_dm(fim)
    cands = []
    for a, b in ((si, sf), (si, fim), (ini, sf)):
        if a is None or b is None or (a, b) == (ini, fim):
            continue
        s = (b - a).days
        if 0 <= s <= MAX_NORMAL_SPAN_DAYS:
            cands.append((s, (a, b)))
    return min(cands, key=lambda c: c[0])[1] if cands else None


def split_list(v) -> list:
    """'A;B' ou uma por linha → ['A', 'B']."""
    if isinstance(v, (list, tuple)):
        parts = [str(x) for x in v]
    else:
        parts = re.split(r"[;\n\r]+", _text(v))
    return [p for p in (_clean_spaces(x) for x in parts) if p]


# ==========================================================================
# Conteúdo programático
# ==========================================================================
def _looks_like_title(t: str) -> bool:
    return t.endswith(":") or bool(_TITLE_RE.match(t))


def _strip_final_dot(t: str) -> str:
    """Remove um ponto final isolado ('Artigo.' → 'Artigo'), preservando 'etc.' e reticências."""
    if t.endswith(".") and not t.endswith("..") and not t.lower().endswith("etc."):
        return t[:-1].rstrip()
    return t


def parse_program(v) -> list:
    """
    Texto da planilha → lista de ProgItem.
      "- item"  → tópico (com marcador)
      "#item"   → título explícito
      sem marcador → título se terminar em ':' ou começar com '1º Dia', 'Módulo 2'…; senão tópico
    Itens vazios (como o ' ' entre dias) são descartados.
    """
    if isinstance(v, (list, tuple)):
        parts = [str(x) for x in v]
    else:
        parts = re.split(r"[;\n\r]+", _text(v))
    items = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if p.startswith("#"):
            t = _clean_spaces(p.lstrip("#"))
            kind = "titulo"
        elif p[0] in _BULLETS:
            t = _clean_spaces(p.lstrip(_BULLETS + " "))
            kind = "topico"
        else:
            t = _clean_spaces(p)
            kind = "titulo" if _looks_like_title(t) else "topico"
        t = _strip_final_dot(t)
        if t:
            items.append(ProgItem(kind, t))
    return items


def program_to_text(items) -> str:
    """Lista de ProgItem → texto no formato da planilha (volta igual em parse_program)."""
    parts = []
    for it in items or []:
        t = _clean_spaces(it.text.replace(";", ","))
        if not t:
            continue
        if it.kind == "titulo":
            parts.append(t if _looks_like_title(t) else "#" + t)
        else:
            parts.append("- " + t)
    return ";".join(parts)


def program_to_df(items) -> pd.DataFrame:
    return pd.DataFrame(
        {"TIPO": [KIND_LABEL[i.kind] for i in items], "TEXTO": [i.text for i in items]}, dtype=object
    )


def df_to_program(df) -> list:
    if df is None or len(df) == 0:
        return []
    items = []
    for _, row in df.iterrows():
        t = _clean_spaces(_text(row.get("TEXTO")))
        if not t:
            continue
        kind = "titulo" if _norm_key(row.get("TIPO")).startswith("tit") else "topico"
        if kind == "topico":
            t = _clean_spaces(t.lstrip(_BULLETS + " ")) or t
        items.append(ProgItem(kind, t))
    return items


def program_sections(items) -> list:
    """
    Blocos selecionáveis para 'conteúdo cursado': cada título com seus tópicos;
    tópicos soltos antes do primeiro título (ou em cursos sem títulos) viram um bloco cada.
    Devolve [(rótulo, [ProgItem, ...]), ...].
    """
    sections, cur = [], None
    for it in items or []:
        if it.kind == "titulo":
            cur = (it.text, [it])
            sections.append(cur)
        elif cur is None:
            sections.append((it.text, [it]))
        else:
            cur[1].append(it)
    return sections


def resolve_parts(text, sections):
    """
    Texto da coluna PARTES → (índices 0-based ordenados, [tokens não entendidos]).
    Aceita números ('1', '1;3', '1,2', '1-3', '1 e 2'), 'todos' e trechos do título ('1º Dia', 'Visita').
    """
    n = len(sections)
    found, bad = set(), []
    raw = _text(text)
    for tok in re.split(r"[;,\n]+|(?<=\d)\s+e\s+(?=\d)", raw):
        tok = _clean_spaces(tok)
        if not tok:
            continue
        key = _norm_key(tok)
        m = re.fullmatch(r"(\d+)\s*(?:-|a|ate)\s*(\d+)", key)
        if key in {"todos", "todas", "tudo", "completo"}:
            found.update(range(n))
        elif m:
            a, b = int(m.group(1)), int(m.group(2))
            if 1 <= a <= b <= n:
                found.update(range(a - 1, b))
            else:
                bad.append(tok)
        elif re.fullmatch(r"\d+", key):
            if 1 <= int(key) <= n:
                found.add(int(key) - 1)
            else:
                bad.append(tok)
        else:
            hits = [i for i, (label, _) in enumerate(sections) if key and key in _norm_key(label)]
            if len(hits) == 1:
                found.add(hits[0])
            else:   # nenhum, ou ambíguo (ex.: 'dia' casa com todos) → pede para ser mais específico
                bad.append(tok)
    return sorted(found), bad


def parts_label(indices) -> str:
    """[0, 1, 3] → '1, 2 e 4' (texto para tabelas)."""
    nums = [str(i + 1) for i in indices]
    if len(nums) <= 1:
        return "".join(nums)
    return ", ".join(nums[:-1]) + " e " + nums[-1]


def _item_keys(items) -> set:
    return {_norm_key(i.text) for i in items or []}


# ==========================================================================
# DataFrames vazios
# ==========================================================================
def empty_cursos_df() -> pd.DataFrame:
    return pd.DataFrame({c: pd.Series(dtype=object) for c in CURSOS_COLS})


def empty_people_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "NOME": pd.Series(dtype=object),
            "TIPO": pd.Series(dtype=object),
            "CURSO": pd.Series(dtype=object),
            "CARGAHORARIA": pd.Series(dtype="float64"),
            "PARTES": pd.Series(dtype=object),
            "CERT_NUM": pd.Series(dtype=object),
            "PROGRAMA": pd.Series(dtype=object),
        }
    )


# ==========================================================================
# Leitura da planilha
# ==========================================================================
_HEADER_ALIASES = {
    "CURSO": {"curso", "nomedocurso", "nomecurso", "evento", "atividade"},
    "DATAINICIO": {"datainicio", "inicio", "datadeinicio", "dtinicio", "datainicial"},
    "DATAFIM": {"datafim", "fim", "datadefim", "dtfim", "datafinal", "termino", "datatermino"},
    "CARGAHORARIA": {"cargahoraria", "carga", "horas", "ch", "cargahorariah"},
    "PROGRAMA": {"programa", "ementa", "conteudo", "conteudoprogramatico", "programacao"},
    "MINISTRANTES": {"ministrantes", "ministrante", "organizadores", "organizador", "equipe", "instrutores"},
    "NOME": {"nome", "nomecompleto", "participante", "participantes", "aluno"},
    "TIPO": {"tipo", "papel", "funcao"},
    "PARTES": {"partes", "parte", "dias", "dia", "modulos", "secoes", "blocos", "cursou", "conteudocursado"},
    "TEXTO": {"texto", "item", "itens", "topico", "topicos", "descricao"},
    "CERT_NUM": {"certnum", "numero", "nrocertificado", "ncertificado", "numerocertificado", "nrcertificado"},
}


def _header_key(s) -> str:
    return re.sub(r"[^a-z0-9]", "", _strip_accents(_text(s)).lower())


def _map_header(row, wanted: set) -> dict:
    """Linha de cabeçalho → {coluna canônica: índice}. Vazio se não achar nenhuma coluna obrigatória."""
    mapping = {}
    for idx, cell in enumerate(row):
        k = _header_key(cell)
        if not k:
            continue
        for canon, aliases in _HEADER_ALIASES.items():
            if canon in wanted and canon not in mapping and k in aliases:
                mapping[canon] = idx
    return mapping


def _read_table(ws, key_col: str, wanted: set):
    """Devolve (mapping, [(nº da linha no Excel, dict coluna→valor)]) ou (None, [])."""
    rows = list(ws.iter_rows(values_only=True))
    for hi, row in enumerate(rows[:6]):
        mapping = _map_header(row, wanted)
        if key_col in mapping:
            data = []
            for ri in range(hi + 1, len(rows)):
                r = rows[ri]
                rec = {c: (r[i] if i < len(r) else None) for c, i in mapping.items()}
                data.append((ri + 1, rec))
            return mapping, data
    return None, []


def _normalize_role_text(t) -> str | None:
    k = _norm_key(t)
    if not k:
        return "participante"
    if k.startswith(("min", "org", "instrut", "palestr", "monitor")):
        return "ministrante"
    if k.startswith(("part", "aluno", "ouvinte", "cursista", "inscrit")):
        return "participante"
    return None


def _norm_certnum(v) -> str:
    t = _text(v)
    if re.fullmatch(r"\d+", t):
        return f"{int(t):04d}"
    return t


def _course_period_notes(nome, ini, fim) -> list:
    notes = []
    if ini and fim:
        if fim < ini:
            sw = swapped_dates(ini, fim)
            extra = f" Seria {sw[0]:%d/%m/%Y} a {sw[1]:%d/%m/%Y}?" if sw else ""
            notes.append(Note("aviso", f"“{nome}”: a data de fim ({fim:%d/%m/%Y}) é anterior ao início ({ini:%d/%m/%Y}).{extra}"))
        else:
            sw = swapped_dates(ini, fim)
            if sw:
                notes.append(
                    Note(
                        "aviso",
                        f"“{nome}”: período de {(fim - ini).days + 1} dias ({ini:%d/%m/%Y} a {fim:%d/%m/%Y}) — "
                        f"dia e mês podem estar trocados. Seria {sw[0]:%d/%m/%Y} a {sw[1]:%d/%m/%Y}? "
                        "(Dica: no Excel, digite as datas em células formatadas como texto.)",
                    )
                )
    return notes


def read_workbook(data: bytes):
    """Lê o .xlsx → (cursos_df, people_df, notes)."""
    notes: list[Note] = []
    wb = load_workbook(io.BytesIO(data), data_only=True)
    sheets = {_header_key(ws.title): ws for ws in wb.worksheets}

    # ---- cursos ---------------------------------------------------------
    cursos_rows = []
    ws_c = sheets.get("cursos")
    if ws_c is None:
        notes.append(Note("aviso", "Aba “cursos” não encontrada — preencha os dados do curso manualmente."))
    else:
        mapping, data_rows = _read_table(ws_c, "CURSO", {"CURSO", "DATAINICIO", "DATAFIM", "CARGAHORARIA", "PROGRAMA", "MINISTRANTES"})
        if mapping is None:
            notes.append(Note("erro", "Na aba “cursos” não achei a coluna CURSO. Use o cabeçalho do modelo."))
        else:
            ignored = 0
            for xl_row, rec in data_rows:
                nome = _clean_spaces(_text(rec.get("CURSO")))
                if not nome:
                    continue
                if nome.upper().startswith(EXAMPLE_PREFIX):
                    ignored += 1
                    continue
                ini, fim = parse_date(rec.get("DATAINICIO")), parse_date(rec.get("DATAFIM"))
                if not _is_blank(rec.get("DATAINICIO")) and ini is None:
                    notes.append(Note("aviso", f"“{nome}”: não entendi a data de início “{_text(rec.get('DATAINICIO'))}” (use dd/mm/aaaa)."))
                if not _is_blank(rec.get("DATAFIM")) and fim is None:
                    notes.append(Note("aviso", f"“{nome}”: não entendi a data de fim “{_text(rec.get('DATAFIM'))}” (use dd/mm/aaaa)."))
                if ini and not fim:
                    fim = ini
                    notes.append(Note("info", f"“{nome}”: sem data de fim — considerei o mesmo dia do início."))
                elif fim and not ini:
                    ini = fim
                carga = parse_hours(rec.get("CARGAHORARIA"))
                if not _is_blank(rec.get("CARGAHORARIA")) and carga is None:
                    notes.append(Note("aviso", f"“{nome}”: não entendi a carga horária “{_text(rec.get('CARGAHORARIA'))}”."))
                notes.extend(_course_period_notes(nome, ini, fim))
                cursos_rows.append(
                    {
                        "CURSO": nome,
                        "DATAINICIO": ini,
                        "DATAFIM": fim,
                        "CARGAHORARIA": carga,
                        "PROGRAMA": program_to_text(parse_program(rec.get("PROGRAMA"))),
                        "MINISTRANTES": ";".join(split_list(rec.get("MINISTRANTES"))),
                    }
                )
            if ignored:
                notes.append(Note("info", f"{ignored} curso(s) de exemplo ignorado(s)."))

    # ---- aba "conteudo": um item por linha (alternativa à célula com ';') -----------
    ws_k = sheets.get("conteudo") or sheets.get("ementa") or sheets.get("ementas")
    if ws_k is not None and cursos_rows:
        mapping_k, data_k = _read_table(ws_k, "TEXTO", {"CURSO", "TIPO", "TEXTO"})
        if mapping_k is None or "CURSO" not in mapping_k:
            notes.append(Note("aviso", "Na aba “conteudo” não achei as colunas CURSO e TEXTO — ignorei a aba."))
        else:
            by_course: dict = {}
            last_course = ""
            for xl_row, rec in data_k:
                texto = _clean_spaces(_text(rec.get("TEXTO")))
                if not texto:
                    continue
                cname = _clean_spaces(_text(rec.get("CURSO"))) or last_course   # curso em branco = o da linha de cima
                last_course = cname
                if not cname or cname.upper().startswith(EXAMPLE_PREFIX):
                    continue
                tipo = _norm_key(rec.get("TIPO"))
                if tipo.startswith("tit"):
                    items = [ProgItem("titulo", _strip_final_dot(texto.lstrip("#").strip()))]
                elif tipo.startswith(("top", "item")):
                    items = [ProgItem("topico", _strip_final_dot(texto.lstrip(_BULLETS + " ").strip()))]
                else:
                    items = parse_program([texto.replace(";", ",")])
                by_course.setdefault(_norm_key(cname), []).extend(items)
            known = {_norm_key(r["CURSO"]) for r in cursos_rows}
            for k in by_course:
                if k not in known:
                    notes.append(Note("aviso", f"Aba “conteudo”: o curso “{k}” não existe na aba “cursos” — confira o nome."))
            for r in cursos_rows:
                items = by_course.get(_norm_key(r["CURSO"]))
                if items:
                    if r["PROGRAMA"]:
                        notes.append(Note("info", f"“{r['CURSO']}”: o conteúdo veio da aba “conteudo” (a coluna PROGRAMA da aba cursos foi ignorada)."))
                    r["PROGRAMA"] = program_to_text(items)
    cursos_df = pd.DataFrame(cursos_rows, columns=CURSOS_COLS, dtype=object) if cursos_rows else empty_cursos_df()

    # ---- pessoas --------------------------------------------------------
    people_rows = []
    ws_p = sheets.get("participantes") or sheets.get("pessoas")
    wanted_p = set(PEOPLE_COLS)
    mapping_p, data_p = (None, [])
    if ws_p is not None:
        mapping_p, data_p = _read_table(ws_p, "NOME", wanted_p)
    else:
        for ws in wb.worksheets:   # planilhas antigas: primeira aba que tenha coluna NOME
            mapping_p, data_p = _read_table(ws, "NOME", wanted_p)
            if mapping_p is not None:
                notes.append(Note("info", f"Aba “participantes” não encontrada — usei a aba “{ws.title}”."))
                break
    if mapping_p is None:
        notes.append(Note("aviso" if ws_p is None else "erro", "Não achei a lista de participantes (coluna NOME)."))
    else:
        ignored = 0
        for xl_row, rec in data_p:
            nome = _clean_spaces(_text(rec.get("NOME")))
            others = any(not _is_blank(rec.get(c)) for c in ("TIPO", "CURSO", "CERT_NUM", "CARGAHORARIA", "PARTES", "PROGRAMA"))
            if not nome:
                if others:
                    notes.append(Note("aviso", f"Linha {xl_row} de participantes tem dados mas está sem nome — ignorada."))
                continue
            if nome.upper().startswith(EXAMPLE_PREFIX):
                ignored += 1
                continue
            role = _normalize_role_text(rec.get("TIPO"))
            if role is None:
                notes.append(Note("aviso", f"Linha {xl_row} ({nome}): tipo “{_text(rec.get('TIPO'))}” não reconhecido — assumi Participante."))
                role = "participante"
            horas = parse_hours(rec.get("CARGAHORARIA"))
            if not _is_blank(rec.get("CARGAHORARIA")) and horas is None:
                notes.append(Note("aviso", f"Linha {xl_row} ({nome}): não entendi a carga horária “{_text(rec.get('CARGAHORARIA'))}”."))
            people_rows.append(
                {
                    "NOME": nome,
                    "TIPO": ROLE_LABEL[role],
                    "CURSO": _clean_spaces(_text(rec.get("CURSO"))) or None,
                    "CARGAHORARIA": float(horas) if horas is not None else float("nan"),
                    "PARTES": _clean_spaces(_text(rec.get("PARTES"))),
                    "CERT_NUM": _norm_certnum(rec.get("CERT_NUM")),
                    "PROGRAMA": program_to_text(parse_program(rec.get("PROGRAMA"))),
                }
            )
        if ignored:
            notes.append(Note("info", f"{ignored} linha(s) de exemplo ignorada(s) na lista de pessoas."))
    if people_rows:
        people_df = pd.DataFrame(people_rows, columns=PEOPLE_COLS)
        people_df["CARGAHORARIA"] = people_df["CARGAHORARIA"].astype("float64")
    else:
        people_df = empty_people_df()

    if len(cursos_df) or len(people_df):
        notes.insert(0, Note("info", f"Planilha lida: {len(cursos_df)} curso(s) e {len(people_df)} pessoa(s)."))
    return cursos_df, people_df, notes


# ==========================================================================
# Pessoas → certificados
# ==========================================================================
def _fem_course(nome: str) -> bool:
    """Heurística de gênero do nome do evento, para 'do/da', 'o/a'."""
    first = _norm_key(nome).split(" ")[0] if _norm_key(nome) else ""
    raw_first = _clean_spaces(nome).split(" ")[0] if nome else ""
    if not first or (raw_first.isupper() and len(raw_first) > 1):   # siglas (PIC, PCI) → masculino
        return False
    if first in {"programa", "sistema", "tema", "problema", "dia", "mapa", "cinema", "diagrama", "esquema", "minicurso"}:
        return False
    if first.endswith(("cao", "sao", "dade", "agem")):
        return True
    return first.endswith("a")


_ARTICLE_FEM = {"do": "da", "o": "a", "no": "na", "ao": "à", "dos": "das", "os": "as", "nos": "nas"}


def role_verb(settings: Settings, role: str, course_nome: str) -> str:
    verb = _clean_spaces(settings.verbos.get(role, "")) or DEFAULT_VERBS[role]
    if _fem_course(course_nome):
        head, _, last = verb.rpartition(" ")
        if last.lower() in _ARTICLE_FEM:
            verb = (head + " " if head else "") + _ARTICLE_FEM[last.lower()]
    return verb


def _fmt_names(names, n=4) -> str:
    names = list(names)
    if len(names) <= n:
        return ", ".join(names)
    return ", ".join(names[:n]) + f" e mais {len(names) - n}"


def build_recipients(course: Course, people_df, settings: Settings, targets: set):
    """Monta a lista final de certificados. Devolve (recs, issues)."""
    issues: list[Note] = []
    pad = settings.padronizar_nomes
    sections = program_sections(course.programa)
    course_key = _norm_key(course.nome)

    # ---- linhas da lista de pessoas --------------------------------------
    rows = []
    other_course = 0
    if people_df is not None:
        for pos, (_, r) in enumerate(people_df.iterrows(), start=1):
            nome = clean_name(r.get("NOME"), pad)
            if not nome:
                if any(not _is_blank(r.get(c)) for c in ("CURSO", "CERT_NUM", "CARGAHORARIA", "PARTES", "PROGRAMA")):
                    issues.append(Note("aviso", f"Linha {pos} da lista de pessoas está sem nome — ignorada."))
                continue
            # coluna CURSO: em branco = vale para o curso selecionado; preenchida = só para aquele curso
            rc = _norm_key(r.get("CURSO"))
            if rc and course_key and rc != course_key:
                other_course += 1
                continue
            role = _normalize_role_text(r.get("TIPO"))
            if role is None:
                issues.append(Note("erro", f"Linha {pos} ({nome}): tipo “{_text(r.get('TIPO'))}” inválido — escolha Participante ou Ministrante."))
                continue
            raw_h = r.get("CARGAHORARIA")
            carga = parse_hours(raw_h)
            if not _is_blank(raw_h) and carga is None:
                issues.append(Note("erro", f"Linha {pos} ({nome}): carga horária “{_text(raw_h)}” inválida."))
                continue
            prog = parse_program(r.get("PROGRAMA"))
            partes = []
            if not _is_blank(r.get("PARTES")):
                idx, bad = resolve_parts(r.get("PARTES"), sections)
                if not sections:
                    issues.append(Note("erro", f"Linha {pos} ({nome}): informou PARTES “{_text(r.get('PARTES'))}”, mas o curso não tem conteúdo cadastrado."))
                    continue
                if bad:
                    issues.append(Note(
                        "erro",
                        f"Linha {pos} ({nome}): não entendi PARTES “{_text(r.get('PARTES'))}” ({', '.join(bad)}). "
                        f"Use números de 1 a {len(sections)} ou parte do título (ex.: 1º Dia).",
                    ))
                    continue
                if prog:
                    issues.append(Note("aviso", f"Linha {pos} ({nome}): PARTES e PROGRAMA preenchidos — usei o PROGRAMA da pessoa."))
                else:
                    partes = idx
                    prog = [it for i in idx for it in sections[i][1]]
            rows.append({"pos": pos, "nome": nome, "role": role, "num": _norm_certnum(r.get("CERT_NUM")), "carga": carga, "prog": prog, "partes": partes})
    if other_course:
        issues.append(Note("info", f"{other_course} linha(s) da lista de pessoas pertencem a outros cursos (coluna CURSO) e foram deixadas de fora."))

    def make(nome, role, row):
        carga = row["carga"] if row else None
        prog = row["prog"] if row else []
        partes = list(row["partes"]) if row else []
        if prog and [(i.kind, _norm_key(i.text)) for i in prog] == [(i.kind, _norm_key(i.text)) for i in course.programa]:
            prog, partes = [], []
        parcial = (carga is not None and carga != course.carga) or bool(prog)
        return Recipient(
            nome=nome, role=role, numero=row["num"] if row else "",
            carga=carga, programa=prog or None, parcial=parcial,
            partes=partes,
        )

    recs: list[Recipient] = []

    # ---- participantes ----------------------------------------------------
    if "participante" in targets:
        seen = {}
        for row in (r for r in rows if r["role"] == "participante"):
            k = _norm_key(row["nome"])
            if k in seen:
                issues.append(Note("aviso", f"Nome repetido entre participantes: {row['nome']} (linhas {seen[k]} e {row['pos']})."))
            else:
                seen[k] = row["pos"]
            recs.append(make(row["nome"], "participante", row))

    # ---- ministrantes: os do curso + ajustes/extras da lista de pessoas ----
    if "ministrante" in targets:
        min_rows = {}
        for row in (r for r in rows if r["role"] == "ministrante"):
            min_rows.setdefault(_norm_key(row["nome"]), row)
        listed = set()
        for nome in course.ministrantes:
            n = clean_name(nome, pad)
            k = _norm_key(n)
            if not n:
                continue
            if k in listed:
                issues.append(Note("aviso", f"Ministrante repetido na lista do curso: {n}."))
                continue
            listed.add(k)
            recs.append(make(n, "ministrante", min_rows.get(k)))
        extras = [r for k, r in min_rows.items() if k not in listed]
        for row in extras:
            recs.append(make(row["nome"], "ministrante", row))
        if extras:
            issues.append(
                Note(
                    "info",
                    f"{_fmt_names([r['nome'] for r in extras])} "
                    f"{'consta como Ministrante' if len(extras) == 1 else 'constam como Ministrantes'} na lista de pessoas, mas não no curso — "
                    f"{'recebe certificado, porém o nome não aparece' if len(extras) == 1 else 'recebem certificado, porém os nomes não aparecem'} "
                    "na lista de organizadores da página 2.",
                )
            )

    # ---- mesma pessoa nos dois papéis ---------------------------------------
    parts = {_norm_key(r.nome) for r in recs if r.role == "participante"}
    both = [r.nome for r in recs if r.role == "ministrante" and _norm_key(r.nome) in parts]
    if both:
        issues.append(Note("aviso", f"{_fmt_names(both)} aparece(m) como participante e ministrante — serão gerados dois certificados."))

    # ---- numeração: manual respeitada, automática pula números já usados ----
    used = {int(r.numero) for r in recs if r.numero.isdigit()}
    counter = int(settings.num_inicial)
    for r in recs:
        if r.numero:
            continue
        while counter in used:
            counter += 1
        r.numero = f"{counter:04d}"
        used.add(counter)
        counter += 1
    return recs, issues


# ==========================================================================
# Validação
# ==========================================================================
def _unsupported_chars(s: str) -> str:
    bad = []
    for ch in s:
        try:
            ch.encode("cp1252")
        except UnicodeEncodeError:
            bad.append(ch)
    return "".join(dict.fromkeys(bad))


def validate(course: Course, recs, settings: Settings, rec_issues, template_ok: bool = True):
    out: list[Note] = list(rec_issues or [])
    add = lambda lvl, msg: out.append(Note(lvl, msg))  # noqa: E731

    if not template_ok:
        add("erro", "Template PDF não encontrado. Envie um PDF na barra lateral.")
    if not course.nome:
        add("erro", "Informe o nome do curso.")
    if course.dt_ini is None or course.dt_fim is None:
        add("erro", "Informe as datas de início e de fim do curso.")
    else:
        if course.dt_fim < course.dt_ini:
            add("erro", "A data de fim é anterior à data de início.")
        sw = swapped_dates(course.dt_ini, course.dt_fim)
        if sw:
            add("aviso", f"As datas do curso parecem ter dia e mês trocados — seria {sw[0]:%d/%m/%Y} a {sw[1]:%d/%m/%Y}?")
        if settings.data_assinatura < course.dt_fim:
            add("aviso", f"A data da assinatura ({settings.data_assinatura:%d/%m/%Y}) é anterior ao fim do curso ({course.dt_fim:%d/%m/%Y}).")
    if not recs:
        add("erro", "Nenhum certificado a gerar — confira a lista de pessoas e o campo “Gerar para”.")

    sem_carga = [r.nome for r in recs if (r.carga if r.carga is not None else course.carga) is None]
    if sem_carga:
        add("erro", "Informe a carga horária do curso" + (f" (faltando para {_fmt_names(sem_carga)})." if course.carga is not None else "."))
    acima = [f"{r.nome} ({format_hours(r.carga)} h)" for r in recs if course.carga and r.carga and r.carga > course.carga]
    if acima:
        add("aviso", f"Carga individual maior que a do curso ({format_hours(course.carga)} h): {_fmt_names(acima)}.")

    # carga × conteúdo: casos em que um lado foi ajustado e o outro não
    so_conteudo = [r.nome for r in recs if r.programa and (r.carga is None or r.carga == course.carga)]
    if so_conteudo:
        add("aviso", f"Conteúdo parcial mas carga horária igual à do curso ({format_hours(course.carga)} h): {_fmt_names(so_conteudo)}. Falta ajustar a carga horária?")
    so_carga = [r.nome for r in recs if r.carga is not None and course.carga and r.carga < course.carga and not r.programa]
    if so_carga:
        add("info", f"Carga horária reduzida mas com o conteúdo completo do curso: {_fmt_names(so_carga)}. Se cursaram só parte, indique as PARTES.")

    # frase
    sample = None
    try:
        sample = _sentence(settings, course, Recipient("Fulano de Tal", "participante"), course.carga or 1)
    except (KeyError, IndexError, ValueError, AttributeError):
        add("erro", "A frase do certificado usa um campo inexistente ou chaves { } inválidas. Campos permitidos: {nome} {papel} {curso} {periodo} {carga}.")
    if sample is not None:
        try:
            Paragraph(sample, _sentence_style(12))
        except Exception:
            add("erro", "A frase do certificado tem marcação inválida — confira se cada <b> tem seu </b>.")

    # conteúdo / organizadores
    if not course.programa:
        add("aviso", "O curso está sem conteúdo programático — a página 2 sairá sem conteúdo.")
    else:
        size, fits = fit_program(tuple(course.programa))
        if not fits:
            add("aviso", "O conteúdo programático completo é grande demais para a página 2, mesmo com fonte reduzida — será cortado. Resuma os itens.")
        elif size < 11:
            add("info", f"O conteúdo programático é longo: a fonte da página 2 será reduzida para {format_hours(size)} pt para caber.")
    for r in recs:
        if r.programa:
            if not fit_program(tuple(r.programa))[1]:
                add("aviso", f"O conteúdo individual de {r.nome} não cabe na página 2.")
    fora = [r.nome for r in recs if r.programa and course.programa and (_item_keys(r.programa) - _item_keys(course.programa))]
    if fora:
        add("info", f"Conteúdo individual com itens que não constam no curso: {_fmt_names(fora)}. Confira se é intencional.")
    if not course.ministrantes:
        add("aviso", "O curso está sem ministrantes — a lista de organizadores da página 2 ficará vazia.")
    elif len(course.ministrantes) > 28:
        add("aviso", "Há mais ministrantes do que cabem na página 2 — parte da lista será cortada.")

    # nomes
    curtos = [r.nome for r in recs if len(r.nome.split()) < 2]
    if curtos:
        add("aviso", f"Nome sem sobrenome: {_fmt_names(curtos)}.")
    ruins = [r.nome for r in recs if any(ch.isdigit() for ch in r.nome)]
    if ruins:
        add("aviso", f"Nome com números: {_fmt_names(ruins)}.")
    chars = _unsupported_chars(" ".join([course.nome] + [r.nome for r in recs] + [i.text for i in course.programa]))
    if chars:
        add("aviso", f"Caracteres que a fonte do certificado não suporta e podem sair errados: {chars}")

    # números
    cont: dict = {}
    for r in recs:
        cont.setdefault(r.numero, []).append(r.nome)
    for num, nomes in cont.items():
        if len(nomes) > 1:
            add("erro", f"Número de certificado {num} repetido: {_fmt_names(nomes)}.")

    # remove mensagens idênticas, mantendo a ordem
    seen, uniq = set(), []
    for n in out:
        if (n.level, n.msg) not in seen:
            seen.add((n.level, n.msg))
            uniq.append(n)
    return uniq


# ==========================================================================
# PDF
# ==========================================================================
A4_W, A4_H = 841.89, 595.28

# Geometria do template padrão (A4 paisagem, origem embaixo à esquerda)
_S1 = dict(x=A4_W * 0.15 + 6, w=A4_W * 0.70 - 12, top=A4_H - 276, h=66)          # frase da pág. 1
_PROG = dict(x=116, w=286, top=A4_H - 127, h=A4_H - 127 - 94)                    # conteúdo (coluna esquerda, pág. 2)
_ORG = dict(cx=576, header_y=A4_H - 163, first_y=A4_H - 177, floor_y=280)        # organizadores (coluna direita)
_REG = dict(x=633, volume_y=218, numero_y=190, max_w=110)                        # caixa "Certificado registrado no livro"
_DATE_Y = 60

_SENT_SIZES = (14, 13, 12, 11)
_PROG_SIZES = (11, 10.5, 10, 9.5, 9, 8.5, 8, 7.5, 7)


def _sentence_style(size) -> ParagraphStyle:
    return ParagraphStyle("cert", fontName="Helvetica", fontSize=size, leading=size * 1.43, alignment=TA_CENTER)


def _sentence(settings: Settings, course: Course, rec: Recipient, carga) -> str:
    if course.dt_ini == course.dt_fim:
        periodo = f"realizado no dia {course.dt_ini:%d/%m/%Y}"
    else:
        periodo = f"realizado de {course.dt_ini:%d/%m/%Y} a {course.dt_fim:%d/%m/%Y}"

    class _Fields(dict):
        def __missing__(self, key):
            raise KeyError(key)

    fields = _Fields(
        nome=_esc(rec.nome),
        papel=_esc(role_verb(settings, rec.role, course.nome)),
        curso=_esc(course.nome),
        periodo=periodo,
        carga=hours_phrase(carga),
    )
    return (settings.frase or DEFAULT_PHRASE).format_map(fields)


def _frame_fits(make_paras, w, h) -> bool:
    c = canvas.Canvas(io.BytesIO(), pagesize=(w + 20, h + 20))
    paras = make_paras()
    Frame(0, 0, w, h, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0).addFromList(paras, c)
    return not paras


def _prog_paragraphs(items, size):
    lead = size * 1.28
    topic = ParagraphStyle(
        "topico", fontName="Helvetica", fontSize=size, leading=lead, leftIndent=size + 3,
        bulletIndent=1, bulletFontName="Helvetica", bulletFontSize=size, spaceBefore=size * 0.22,
    )
    title = ParagraphStyle(
        "titulo", fontName="Helvetica-Bold", fontSize=size, leading=lead,
        spaceBefore=size * 0.9, spaceAfter=size * 0.1,
    )
    out = []
    for it in items:
        if it.kind == "titulo":
            out.append(Paragraph(_esc(it.text), title))
        else:
            out.append(Paragraph(_esc(it.text), topic, bulletText="•"))
    return out


@lru_cache(maxsize=512)
def fit_program(items: tuple):
    """(tamanho_de_fonte, coube?) para o conteúdo programático na coluna esquerda da página 2."""
    if not items:
        return _PROG_SIZES[0], True
    for s in _PROG_SIZES:
        if _frame_fits(lambda: _prog_paragraphs(items, s), _PROG["w"], _PROG["h"]):
            return s, True
    return _PROG_SIZES[-1], False


@lru_cache(maxsize=512)
def _fit_sentence(html_text: str):
    for s in _SENT_SIZES:
        if _frame_fits(lambda: [Paragraph(html_text, _sentence_style(s))], _S1["w"], _S1["h"]):
            return s, True
    return _SENT_SIZES[-1], False


def _fit_width(text, font, size, max_w, min_size=5.5):
    while size > min_size and stringWidth(text, font, size) > max_w:
        size -= 0.5
    return size


def _names_layout(n: int):
    """Escolhe colunas/fonte para n organizadores. Devolve (colunas, tamanho, espaçamento, cabem)."""
    avail = _ORG["first_y"] - _ORG["floor_y"]
    for cols, sizes in ((1, (8, 7.5, 7)), (2, (7, 6.5, 6))):
        per_col = math.ceil(n / cols)
        for s in sizes:
            lead = s * 1.5
            if (per_col - 1) * lead <= avail:
                return cols, s, lead, True
    s, lead = 6, 9
    return 2, s, lead, False


def _overlay_p1(W, H, sentence_html, date_line):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(W, H))
    c.scale(W / A4_W, H / A4_H)
    size, fits = _fit_sentence(sentence_html)
    Frame(
        _S1["x"], _S1["top"] - _S1["h"], _S1["w"], _S1["h"],
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
    ).addFromList([Paragraph(sentence_html, _sentence_style(size))], c)
    c.setFont("Helvetica", 12)
    c.drawCentredString(A4_W / 2, _DATE_Y, date_line)
    c.save()
    packet.seek(0)
    return packet, fits


def _overlay_p2(W, H, items, ministrantes, volume, numero, date_line):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(W, H))
    c.scale(W / A4_W, H / A4_H)
    notes = []

    # conteúdo programático
    size, fits = fit_program(tuple(items))
    if items:
        Frame(
            _PROG["x"], _PROG["top"] - _PROG["h"], _PROG["w"], _PROG["h"],
            leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        ).addFromList(_prog_paragraphs(items, size), c)
        if not fits:
            notes.append("o conteúdo programático não coube na página 2 e foi cortado")

    # organizadores
    cx = _ORG["cx"]
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(cx, _ORG["header_y"], "Organizadores:")
    if ministrantes:
        cols, s, lead, ok = _names_layout(len(ministrantes))
        per_col = math.ceil(len(ministrantes) / cols)
        centers = [cx] if cols == 1 else [cx - 80, cx + 80]
        max_w = 300 if cols == 1 else 150
        shown = 0
        for i, nome in enumerate(ministrantes):
            col, row = divmod(i, per_col)
            y = _ORG["first_y"] - row * lead
            if y < _ORG["floor_y"] - 0.5:
                continue
            fs = _fit_width(nome, "Helvetica", s, max_w)
            c.setFont("Helvetica", fs)
            c.drawCentredString(centers[min(col, cols - 1)], y, nome)
            shown += 1
        if shown < len(ministrantes):
            notes.append(f"só {shown} de {len(ministrantes)} organizadores couberam na página 2")

    # caixa de registro
    if volume:
        c.setFont("Helvetica", _fit_width(volume, "Helvetica", 12, _REG["max_w"]))
        c.drawCentredString(_REG["x"], _REG["volume_y"], volume)
    c.setFont("Helvetica", _fit_width(numero, "Helvetica", 12, _REG["max_w"]))
    c.drawCentredString(_REG["x"], _REG["numero_y"], numero)

    c.setFont("Helvetica", 12)
    c.drawCentredString(A4_W / 2, _DATE_Y, date_line)
    c.save()
    packet.seek(0)
    return packet, notes


def render_certificate(template_bytes: bytes, course: Course, rec: Recipient, settings: Settings):
    """Gera o PDF de uma pessoa. Devolve (pdf_bytes, [avisos])."""
    notes: list[str] = []
    reader = PdfReader(io.BytesIO(template_bytes))
    carga = rec.carga if rec.carga is not None else course.carga
    date_line = f"{settings.local}, {settings.data_assinatura:%d/%m/%Y}"
    writer = PdfWriter()

    p1 = reader.pages[0]
    W, H = float(p1.mediabox.width), float(p1.mediabox.height)
    if abs(W - A4_W) > 3 or abs(H - A4_H) > 3:
        notes.append("o template não é A4 paisagem — as posições dos textos podem ficar deslocadas")
    sentence = _sentence(settings, course, rec, carga)
    packet1, fits1 = _overlay_p1(W, H, sentence, date_line)
    if not fits1:
        notes.append("a frase da página 1 ficou longa demais e foi cortada")
    p1.merge_page(PdfReader(packet1).pages[0])
    writer.add_page(p1)

    if len(reader.pages) > 1:
        p2 = reader.pages[1]
        items = rec.programa if rec.programa else course.programa
        mins = [clean_name(m, settings.padronizar_nomes) for m in course.ministrantes]
        packet2, n2 = _overlay_p2(float(p2.mediabox.width), float(p2.mediabox.height), items, mins, settings.volume, rec.numero, date_line)
        notes.extend(n2)
        p2.merge_page(PdfReader(packet2).pages[0])
        writer.add_page(p2)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue(), notes


def _safe_filename(s: str) -> str:
    s = re.sub(r'[\\/:*?"<>|]+', "-", s)
    return _clean_spaces(s).strip(" .") or "sem-nome"


def build_zip(template_bytes, course: Course, recs, settings: Settings, progress=None):
    """Gera todos os PDFs num ZIP (+ _registro_certificados.csv). Devolve (zip_bytes, [avisos])."""
    buf = io.BytesIO()
    note_map: dict = {}
    failures: list[str] = []
    registro = [["numero", "nome", "tipo", "carga_horaria", "conteudo", "partes", "arquivo"]]
    used_paths = set()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for i, r in enumerate(recs):
            try:
                pdf, rn = render_certificate(template_bytes, course, r, settings)
            except Exception as e:  # um nome problemático não derruba o lote
                failures.append(f"{FAIL_PREFIX}{r.nome} ({e})")
            else:
                folder = "Participantes" if r.role == "participante" else "Ministrantes"
                base = f"{folder}/{_safe_filename(r.numero)} - {_safe_filename(r.nome)}"
                path, k = base + ".pdf", 2
                while path in used_paths:
                    path, k = f"{base} ({k}).pdf", k + 1
                used_paths.add(path)
                z.writestr(path, pdf)
                for n in rn:
                    note_map.setdefault(n, []).append(r.nome)
                carga = r.carga if r.carga is not None else course.carga
                registro.append([r.numero, r.nome, ROLE_LABEL[r.role], format_hours(carga), "Parcial" if r.programa else "Completo", parts_label(r.partes), path])
            if progress:
                progress((i + 1) / len(recs))
        sio = io.StringIO()
        csv.writer(sio, delimiter=";").writerows(registro)
        z.writestr("_registro_certificados.csv", ("\ufeff" + sio.getvalue()).encode("utf-8"))
    if recs and len(failures) == len(recs):
        raise RuntimeError(failures[0])
    notes = failures + [f"{_fmt_names(names, 5)}: {msg}." for msg, names in note_map.items()]
    return buf.getvalue(), notes


# ==========================================================================
# Planilha modelo / exportação
# ==========================================================================
_NAVY = "1B2A4A"
_THIN = Side(style="thin", color="D0D6E2")


def _style_header(ws, ncols):
    for col in range(1, ncols + 1):
        c = ws.cell(row=1, column=col)
        c.font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
        c.fill = PatternFill("solid", fgColor=_NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(bottom=Side(style="medium", color="C8782A"))
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"


def _validation(ws, rng, prompt_title, prompt, error=None, **kw):
    dv = DataValidation(
        allow_blank=True, showInputMessage=True, promptTitle=prompt_title[:32], prompt=prompt[:255],
        showErrorMessage=error is not None, errorTitle="Valor inválido" if error else None, error=error, **kw,
    )
    dv.add(rng)
    ws.add_data_validation(dv)


def _instructions_sheet(ws):
    lines = [
        ("Como preencher a planilha do Gerador de Certificados", "title"),
        ("", None),
        ("RESUMO: preencha só o que for diferente do normal. Células em branco usam o padrão do curso.", "h"),
        ("", None),
        ("ABA “cursos” — uma linha por curso", "h"),
        ("• CURSO: tipo e nome (ex.: Minicurso de LaTeX, Oficina de PCI).", None),
        ("• DATAINICIO / DATAFIM: dd/mm/aaaa (ex.: 21/09/2026). Curso de um dia só: repita a data. O site avisa se dia e mês parecerem trocados.", None),
        ("• CARGAHORARIA: horas do curso completo (30, 2,5 ou 2h30).", None),
        ("• MINISTRANTES: nomes separados por ponto e vírgula (;). Eles recebem certificado de ministrante automaticamente.", None),
        ("• PROGRAMA (opcional): o conteúdo em uma célula só, separado por “;”. Se preferir, deixe em branco e use a aba “conteudo”, que é mais fácil.", None),
        ("", None),
        ("ABA “conteudo” — um item por linha (mais fácil que a célula com “;”)", "h"),
        ("• CURSO: escolha na lista (a linha de baixo pode ficar em branco para repetir o curso de cima).", None),
        ("• TIPO: Título (aparece em negrito, ex.: “1º Dia: Fundamentos”) ou Tópico (aparece com marcador).", None),
        ("• TEXTO: o item, com a primeira letra maiúscula e sem ponto final.", None),
        ("• Cada Título com seus Tópicos forma uma PARTE numerada (1, 2, 3…), usada para participação parcial.", None),
        ("", None),
        ("ABA “participantes” — uma linha por pessoa", "h"),
        ("• NOME: nome completo (pode estar em MAIÚSCULAS; o sistema padroniza).", None),
        ("• TIPO: Participante ou Ministrante (lista suspensa). Ministrantes do curso já entram sozinhos — só liste aqui para ajuste individual.", None),
        ("• CURSO: escolha o curso da pessoa. Com vários cursos na mesma planilha, isso evita misturar as listas. Em branco = vale para o curso aberto no site.", None),
        ("", None),
        ("CASOS ESPECÍFICOS (deixe em branco quando não se aplica)", "h"),
        ("• Carga horária parcial: preencha CARGAHORARIA da pessoa (ex.: 8 em um curso de 30 h).", None),
        ("• Conteúdo parcial: preencha PARTES com os números das partes que a pessoa cursou. Ex.: 1  ·  1;3  ·  1-2  ·  ou parte do título, como “1º Dia”.", None),
        ("   (Se preferir, PROGRAMA aceita o texto do conteúdo, no mesmo formato da aba cursos — mas PARTES é bem mais simples.)", None),
        ("• Número fixo: preencha CERT_NUM só para fixar um número (confira o Livro de Registros). Em branco = numeração automática.", None),
        ("• Linhas de exemplo começam com “EXEMPLO” e são ignoradas — pode apagá-las.", None),
        ("", None),
        ("Dica: no site você também pode preencher tudo sem planilha e baixar a planilha atualizada no final.", None),
    ]
    ws.column_dimensions["A"].width = 130
    for i, (txt, kind) in enumerate(lines, start=1):
        c = ws.cell(row=i, column=1, value=txt)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if kind == "title":
            c.font = Font(bold=True, size=15, color=_NAVY)
        elif kind == "h":
            c.font = Font(bold=True, size=12, color="C8782A")


def _build_workbook(cursos_rows, content_rows, people_rows, example: bool) -> bytes:
    wb = Workbook()
    ws_i = wb.active
    ws_i.title = "instruções"
    _instructions_sheet(ws_i)
    gray = Font(italic=True, color="7A8295")
    course_list = "cursos!$A$2:$A$60"

    # ---- cursos ----------------------------------------------------------
    ws = wb.create_sheet("cursos")
    ws.append(CURSOS_COLS)
    _style_header(ws, len(CURSOS_COLS))
    for width, col in zip((34, 14, 14, 14, 78, 62), "ABCDEF"):
        ws.column_dimensions[col].width = width
    for r in cursos_rows:
        ws.append(r)
    last = max(len(cursos_rows) + 1, 60)
    for row in ws.iter_rows(min_row=2, max_row=last):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if example and c.row <= len(cursos_rows) + 1 and c.value is not None:
                c.font = gray
        row[1].number_format = row[2].number_format = "@" if example else "dd/mm/yyyy"
        row[3].number_format = "General"
    _validation(ws, f"B2:C{last}", "Data (dd/mm/aaaa)", "Digite como 21/09/2026. Curso de um dia: repita a data.",
                "Use o formato dd/mm/aaaa.", type="textLength", operator="lessThanOrEqual", formula1="10")
    _validation(ws, f"D2:D{last}", "Carga horária", "Horas do curso completo, ex.: 30 ou 2,5.",
                "Digite um número maior que zero (ex.: 30 ou 2,5).", type="decimal", operator="greaterThan", formula1="0")
    _validation(ws, f"E2:E{last}", "Conteúdo (opcional aqui)",
                "Prefira a aba “conteudo” (um item por linha). Aqui: itens separados por “;”, “-” para tópicos, sem ponto final.",
                type="textLength", operator="lessThanOrEqual", formula1="8000")
    _validation(ws, f"F2:F{last}", "Ministrantes", "Nomes separados por “;”, sem ponto no final.", type="textLength", operator="lessThanOrEqual", formula1="4000")

    # ---- conteudo ----------------------------------------------------------
    wk = wb.create_sheet("conteudo")
    wk.append(CONTENT_COLS)
    _style_header(wk, len(CONTENT_COLS))
    for width, col in zip((34, 12, 90), "ABC"):
        wk.column_dimensions[col].width = width
    for r in content_rows:
        wk.append(r)
    lastk = max(len(content_rows) + 1, 400)
    for row in wk.iter_rows(min_row=2, max_row=lastk):
        row[2].alignment = Alignment(wrap_text=True, vertical="top")
        if example and row[0].row <= len(content_rows) + 1 and row[2].value is not None:
            for c in row:
                c.font = gray
    _validation(wk, f"A2:A{lastk}", "Curso", "Escolha o curso (ou deixe em branco para repetir o de cima).",
                "Esse curso não está na aba “cursos”. Confira o nome.", type="list", formula1=course_list, errorStyle="warning")
    _validation(wk, f"B2:B{lastk}", "Tipo", "Título = negrito (ex.: 1º Dia: Fundamentos). Tópico = com marcador. Em branco = Tópico.",
                "Escolha Título ou Tópico.", type="list", formula1='"Título,Tópico"')
    _validation(wk, f"C2:C{lastk}", "Texto do item", "Primeira letra maiúscula, sem ponto final.")

    # ---- participantes -----------------------------------------------------
    wp = wb.create_sheet("participantes")
    wp.append(PEOPLE_COLS)
    _style_header(wp, len(PEOPLE_COLS))
    for width, col in zip((42, 15, 34, 16, 14, 14, 50), "ABCDEFG"):
        wp.column_dimensions[col].width = width
    for r in people_rows:
        wp.append(r)
    lastp = max(len(people_rows) + 1, 300)
    for row in wp.iter_rows(min_row=2, max_row=lastp):
        row[4].number_format = "@"        # PARTES: "1;3" não vira data/número
        row[5].number_format = "@"        # CERT_NUM
        row[3].number_format = "General"
        row[6].alignment = Alignment(wrap_text=True, vertical="top")
        if example and row[0].row <= len(people_rows) + 1 and row[0].value is not None:
            for c in row:
                c.font = gray
    _validation(wp, f"B2:B{lastp}", "Tipo", "Participante ou Ministrante.", "Escolha Participante ou Ministrante.",
                type="list", formula1='"Participante,Ministrante"')
    _validation(wp, f"C2:C{lastp}", "Curso", "Escolha o curso da pessoa. Em branco = curso aberto no site.",
                "Esse curso não está na aba “cursos”. Confira o nome.", type="list", formula1=course_list, errorStyle="warning")
    _validation(wp, f"D2:D{lastp}", "Carga horária (opcional)", "Em branco = carga horária do curso. Preencha para participação parcial (ex.: 8).",
                "Digite um número maior que zero (ex.: 8 ou 2,5).", type="decimal", operator="greaterThan", formula1="0")
    _validation(wp, f"E2:E{lastp}", "Partes cursadas (opcional)",
                "Em branco = conteúdo completo. Números das partes do curso: 1  ·  1;3  ·  1-2  ·  ou parte do título (1º Dia).")
    _validation(wp, f"F2:F{lastp}", "Nº do certificado", "Em branco = numeração automática. Preencha só para fixar um número.")
    _validation(wp, f"G2:G{lastp}", "Conteúdo por texto (opcional)",
                "Alternativa a PARTES: o conteúdo cursado em texto, itens separados por “;”.", type="textLength", operator="lessThanOrEqual", formula1="8000")
    wb.active = 3
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _content_rows(cursos_rows_named):
    """[(nome_do_curso, programa_texto)] → linhas da aba 'conteudo' (curso só na primeira linha de cada bloco)."""
    rows = []
    for nome, prog_text in cursos_rows_named:
        first = True
        for it in parse_program(prog_text):
            rows.append([nome if first else None, KIND_LABEL[it.kind], it.text])
            first = False
    return rows


def make_template_workbook() -> bytes:
    exemplo = f"{EXAMPLE_PREFIX} – Minicurso de LaTeX"
    prog = (
        "1º Dia: Fundamentos;- Introdução ao LaTeX e Overleaf;- Equações e modo matemático;"
        "2º Dia: Elementos gráficos;- Imagens e tabelas;- Introdução à plotagem de gráficos"
    )
    cursos = [[exemplo, "20/06/2026", "27/06/2026", 30, None, "Maria da Silva;João Pereira"]]
    people = [
        # NOME, TIPO, CURSO, CARGAHORARIA, PARTES, CERT_NUM, PROGRAMA
        [f"{EXAMPLE_PREFIX} – Ana Souza", "Participante", exemplo, None, None, None, None],
        [f"{EXAMPLE_PREFIX} – Bruno Lima (só 8 h)", "Participante", exemplo, 8, None, None, None],
        [f"{EXAMPLE_PREFIX} – Carla Dias (só o 1º dia)", "Participante", exemplo, 15, "1", None, None],
        [f"{EXAMPLE_PREFIX} – Diego Alves (os dois dias, número fixo)", "Participante", exemplo, None, "1;2", "0150", None],
    ]
    return _build_workbook(cursos, _content_rows([(exemplo, prog)]), people, example=True)


def export_workbook(cursos_df, people_df) -> bytes:
    cursos, named = [], []
    for _, r in (cursos_df if cursos_df is not None else empty_cursos_df()).iterrows():
        nome = _clean_spaces(_text(r.get("CURSO")))
        if not nome:
            continue
        carga = parse_hours(r.get("CARGAHORARIA"))
        cursos.append(
            [nome, parse_date(r.get("DATAINICIO")), parse_date(r.get("DATAFIM")), carga if carga is not None else None,
             None, ";".join(split_list(r.get("MINISTRANTES")))]
        )
        named.append((nome, _text(r.get("PROGRAMA"))))
    people = []
    for _, r in (people_df if people_df is not None else empty_people_df()).iterrows():
        nome = _clean_spaces(_text(r.get("NOME")))
        if not nome:
            continue
        role = _normalize_role_text(r.get("TIPO")) or "participante"
        people.append(
            [nome, ROLE_LABEL[role], _clean_spaces(_text(r.get("CURSO"))) or None, parse_hours(r.get("CARGAHORARIA")),
             _clean_spaces(_text(r.get("PARTES"))) or None, _norm_certnum(r.get("CERT_NUM")) or None,
             program_to_text(parse_program(r.get("PROGRAMA"))) or None]
        )
    return _build_workbook(cursos, _content_rows(named), people, example=False)


def people_without_course(people_df, cursos_df) -> list:
    """Nomes cuja coluna CURSO não bate com nenhum curso da planilha (provável erro de digitação)."""
    if people_df is None or len(people_df) == 0:
        return []
    known = {_norm_key(c) for c in (cursos_df["CURSO"] if cursos_df is not None and len(cursos_df) else [])}
    bad = []
    for _, r in people_df.iterrows():
        k = _norm_key(r.get("CURSO"))
        if k and known and k not in known:
            bad.append(f"{clean_name(r.get('NOME'), False)} (“{_text(r.get('CURSO'))}”)")
    return bad
