from __future__ import annotations
import re
import unicodedata
from typing import Iterable, Set, Dict, Optional
import spacy

def _get_nlp():
    try:
        return spacy.load("es_core_news_sm")
    except OSError:
        # Modelo no instalado: lo descargamos y cargamos
        from spacy.cli import download
        download("es_core_news_sm")
        return spacy.load("es_core_news_sm")


# --- Add this helper inside Titulos.py ---
def _bold_before_pipe(text, shape):
    """
    Applies bold formatting to all text before and including the first '|' in the shape.
    """
    try:
        if "|" not in text:
            return
        before, sep, after = text.partition("|")
        rng = shape.TextFrame.TextRange
        full_text = before + sep + after
        rng.Text = full_text
        # Bold up to and including the '|'
        rng.Characters(1, len(before) + len(sep)).Font.Bold = True
        # Unbold the rest
        rng.Characters(len(before) + len(sep) + 1, len(full_text) - len(before) - len(sep)).Font.Bold = False
    except Exception as e:
        print(f"Bold formatting failed: {e}")


# --- Opcional spaCy ---
_USE_SPACY = True
try:
    import spacy
    _NLP = None
except Exception:
    spacy = None
    _USE_SPACY = False
    _NLP = None

def _lazy_nlp():
    global _NLP, _USE_SPACY
    if not _USE_SPACY:
        return None
    if _NLP is None:
        try:
            _NLP = spacy.load("es_core_news_sm")
        except Exception:
            # sin internet o sin permisos -> no bloqueamos; seguimos sin spacy
            _USE_SPACY = False
            _NLP = None
    return _NLP

def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

# Equivalencias de frases
_PHRASE_EQUIV: Dict[str, str] = {
    "ns/nc": "no sabe",
    "no lo se": "no sabe",
    "no se": "no sabe",
    "no lo sabe": "no sabe",
    "no sabe / no contesta": "no sabe",
    "no sabe/ no contesta": "no sabe",
    "nunca lo votaria": "nunca lo votaría",
}

# Equivalencias rápidas de tokens (género/número)
_TOKEN_EQUIV: Dict[str, str] = {
    "buena": "bueno", "buenas": "bueno", "buenos": "bueno",
    "mala": "malo", "malas": "malo", "malos": "malo",
    "informada": "informado", "informadas": "informado", "informados": "informado",
    "corrupta": "corrupto", "corruptas": "corrupto", "corruptos": "corrupto",
}

_ALLOWED_EXTRAS: Set[str] = set()

def _normalize_answer_spacy(ans: str) -> str:
    nlp = _lazy_nlp()
    if nlp is None:
        return _normalize_answer_plain(ans)

    a = ans.strip().lower()
    a = _strip_accents(a)
    a = re.sub(r"\s+", " ", a)

    if a in _PHRASE_EQUIV:
        a = _PHRASE_EQUIV[a]

    doc = nlp(a)
    toks = []
    for t in doc:
        if t.is_space or t.is_punct:
            continue
        lemma = t.lemma_.lower()
        lemma = _TOKEN_EQUIV.get(lemma, lemma)
        toks.append(lemma)

    norm = " ".join(toks)
    if norm == "nunca lo votaria":
        norm = "nunca lo votaría"
    return norm.strip()

def _normalize_answer_plain(ans: str) -> str:
    a = ans.strip().lower()
    a = _strip_accents(a)
    a = re.sub(r"\s+", " ", a)
    if a in _PHRASE_EQUIV:
        a = _PHRASE_EQUIV[a]
    # Normalizaciones sencillas por palabra
    words = [ _TOKEN_EQUIV.get(w, w) for w in a.split() ]
    norm = " ".join(words)
    if norm == "nunca lo votaria":
        norm = "nunca lo votaría"
    return norm.strip()

def _normalize_answer(ans: str) -> str:
    return _normalize_answer_spacy(ans) if _USE_SPACY else _normalize_answer_plain(ans)

def _norm_set(items: Iterable[str]) -> Set[str]:
    return { _normalize_answer(x) for x in items if x is not None and str(x).strip() != "" }

def _canon_sets() -> Dict[str, Set[str]]:
    ordinal_val_5 = _norm_set(["muy buena", "buena", "mala", "muy mala", "no sabe"])
    ordinal_val_7 = _norm_set(["muy bien", "bien", "positivo regular", "negativo regular", "malo", "muy malo", "no lo se"])
    ord5_corrupto    = _norm_set(["muy corrupto","bastante corrupto","poco corrupto","nada corrupto","no sabe"])
    ord5_resp        = _norm_set(["muy responsable","bastante responsable","poco responsable","nada responsable","no sabe"])
    ord5_info        = _norm_set(["muy informado","bastante informado","poco informado","nada informado","no sabe"])
    ord5_grave       = _norm_set(["muy grave","bastante grave","poco grave","nada grave","no sabe"])
    ord5_acuerdo     = _norm_set(["muy de acuerdo","de acuerdo","en desacuerdo","muy en desacuerdo","no sabe"])
    ord5_contcambio  = _norm_set([
        "continuar como hasta ahora","continuar con algunos cambios",
        "cambiar manteniendo solo algunas cosas","cambiar totalmente","no sabe"
    ])
    potencialidad = _norm_set(["muy probable","bastante probable","poco probable","nunca lo votaria","no sabe"] or ["muy probable","bastante probable","poco probable","nunca lo votaria","No lo sé aun"] or ['Muy probable', 'Bastante probable', 'No sabe', 'Poco probable', 'Nada probable'])
    clase_social = _norm_set(["Clase baja","Clase media baja","Clase media","Clase media alta","Clase alta","No sabe"])
    motivo_voto  = _norm_set([
        "A favor de Milei","En contra de Milei",
        "A favor del kirchnerismo","En contra del kirchnerismo",
        "A favor de la lista de Federico Achával","En contra de la lista de Federico Achával",
        "No sabe"
    ])
    afirmativa = _norm_set(["si","no","no sabe"])

    return {
        "ordinal_valorativa_5": ordinal_val_5,
        "ordinal_valorativa_7": ordinal_val_7,
        "ordinal_5": ord5_corrupto,
        "ordinal_5": ord5_resp,
        "ordinal_5": ord5_info,
        "ordinal_5": ord5_grave,
        "ordinal_5": ord5_acuerdo,
        "ordinal_5": ord5_contcambio,
        "potencialidad": potencialidad,
        "clase_social": clase_social,
        "principal_motivo_del_voto": motivo_voto,
        "afirmativa": afirmativa,
    }

def clasificar_tipo_spacy(respuestas: Iterable[str], nombre: Optional[str] = None) -> str:
    canons = _canon_sets()
    resp_set = _norm_set(respuestas) - _ALLOWED_EXTRAS

    # 1) match exacto
    for tipo, canon in canons.items():
        if resp_set == canon:
            return tipo

    # 2) match permisivo: subset/superset con mínima diferencia
    def symdiff(a: Set[str], b: Set[str]) -> int:
        return len(a.symmetric_difference(b))
    candidatos = []
    for tipo, canon in canons.items():
        if resp_set.issubset(canon) or canon.issubset(resp_set):
            candidatos.append((tipo, symdiff(resp_set, canon)))
    if candidatos:
        candidatos.sort(key=lambda x: x[1])
        return candidatos[0][0]

    # 3) si nada encaja → nominal
    return "nominal"

# API compatible con tu código existente
def Determinar_tipo(respuestas, nombre):
    return clasificar_tipo(respuestas, nombre)
