from dataclasses import dataclass, field
from typing import List, Optional, Any
import hashlib, unicodedata

class question:
    # ---- Source data (as parsed) ----
    nombre_raw: str                      # La pregunta original (texto crudo)
    respuestas: List[str] = field(default_factory=list)   # Lista de respuestas
    bloque_df: Any = None                # El bloque completo (DataFrame u otro)
    ubicacion: Optional[str] = None      # Hoja/rango original o pista de ubicación
    tipo: Optional[str] = None           # Tipo inferido

    # ---- Manifest fields (single source of truth) ----
    id: str = field(init=False)                          # Hash estable
    titulo: Optional[str] = None                         # Output de determinar_titulo
    excel_range_block: Optional[str] = None              # A1 del bloque (e.g. Sheet!$B$2:$BS$7)
    excel_ranges_per_row: List[str] = field(default_factory=list)  # A1 por fila
    legends_original: List[str] = field(default_factory=list)      # Etiquetas tal cual
    legends_norm: List[str] = field(default_factory=list)          # Etiquetas normalizadas
    palette_hint: Optional[str] = None                   # Sugerencia de paleta (COLOR_MAP)

    # ---- Placeholders para mapping PPT ----
    slide_index: Optional[int] = None
    slide_id: Optional[str] = None
    chart_rel_id: Optional[str] = None
    chart_xml_path: Optional[str] = None

    # ---- Compatibilidad hacia atrás ----
    nombre: Optional[str] = None  # puedes seguir accediendo como antes

#tipos de pregunta

pregunta_valorativa_fem = ["muy buena", "buena", "mala", "muy mala", "no sabe"]
pregunta_valorativa_masc = ["muy bueno", "bueno", "malo", "muy malo", "no sabe"]
pregunta_positivo_o_Regular = ["muy bien", "bien", "positivo regular", "negativo regular", "malo", "muy malo", "no lo se"]
pregunta_afirmativa = ["si", "no", "no sabe"]
pregunta_grado_de_acuerdo = ["muy de acuerdo", "de acuerdo", "en desacuerdo", "muy en desacuerdo", "No sabe"]
pregunta_potencialidad_de_voto_masc = ["muy probable", "bastante probable", "no sabe", "poco probable", "nunca lo votaria"]
pregunta_potencialidad_de_voto_fem = ["muy probable", "bastante probable", "no sabe", "poco probable", "nunca la votaria"]
pregunta_de_conocimiento_masc = ["muy buena", "buena", "mala", "muy mala", "no sabe", "no lo conoce o no tiene opinion", "No sabe"]
pregunta_de_conocimiento_fem = ["muy buena", "buena", "mala", "muy mala", "no sabe", "no lo conoce o no tiene opinion", "No sabe"]

################# DICCIONARIOS ########################

ANSWER_SETS = {
    "valorativa_5": {"muy buena", "buena", "mala", "muy mala", "no sabe"},
    "valorativa_7": {"muy bien", "bien", "positivo regular", "negativo regular", "malo", "muy malo", "no lo se"},
    "afirmativa": {"si", "no", "no sabe"},
    "acuerdo": {"muy de acuerdo", "de acuerdo", "en desacuerdo", "muy en desacuerdo", "no sabe"},
    "potencialidad": {"muy probable", "bastante probable", "poco probable", "nunca lo votaria", "no sabe"},
    "conocimiento": {"muy buena", "buena", "mala", "muy mala", "no sabe", "no lo conoce o no tiene opinion"},
    "CoC":{"continuar como hasta ahora", "continuar con algunos cambios", "cambiar manteniendo solo algunas cosas", "cambiar totalmente", "no sabe"},
    "grado de informacion": {'muy informado', 'bastante informado', 'poco informado', 'nada informado', 'no sabe'},
    "nivel de gravedad": {'muy grave', 'bastante grave', 'poco grave', 'nada grave', 'no sabe'},
    "cantidad": {'Mucho', 'Bastante', 'Poco', 'Nada', 'No sabe'},
    "nivel de responsabilidad": {'muy responsable', 'bastante responsable', 'poco responsable', 'nada responsable', 'no sabe'},
    "nivel de corrupccion": {'muy corrupto', 'bastante corrupto', 'poco corrupto', 'nada corrupto', 'no sabe'}
}

TITLE_CUES = {
    "imagen": ["imagen"],
    "gestion": ["como calificas", "gestion"],
    "multiple": ["multiple"],
    "clase social": ["clase social"],
    "estimulo": ["video", "spot", "SPOT"],
    "afirmacion": ["afirmacion"],
    "principal motivo del voto": ["motivo"]
}

COLOR_MAP = {
    # categorical / “nominal”
    "nominal": 1,

    # valorativas (gestión / imagen / estímulo): verde→rojo (+ gris)
    "ordinal_valorativa_5": 2,
    "valorativa_5": 2,              # alias, por si tu parser varía
    "ordinal_valorativa_7": 4,
    "valorativa_7": 4,              # alias

    # ordinales de 5 puntos (acuerdo, gravedad, responsabilidad, corrupción, coc, info…)
    "ordinal_5": 3,
    "cantidad": 3,

    # binarias / afirmativas
    "afirmativa": 6,

    # intención / potencialidad de voto
    "potencialidad": 5,

    # clase social
    "clase_social": 6,

    # múltiple selección
    "multiple": 1,

    # motivo del voto
    "principal_motivo_del_voto": 7,
}


#ordinales valorativas: gestion estimulo imagen
#ordinales: afirmacion conocimiento gravedad responsabilidad corrupción acuerdo continuidad o cambio
#Ordinal_valorativa_5: ["muy buena", "buena", "mala", "muy mala", "no sabe"]
#Ordinal_valorativa_7: ["muy bien", "bien", "positivo regular", "negativo regular", "malo", "muy malo", "no lo se"]
#Ordinal_5: {'muy corrupto', 'bastante corrupto', 'poco corrupto', 'nada corrupto', 'no sabe'}{'muy responsable', 'bastante responsable', 'poco responsable', 'nada responsable', 'no sabe'}{"muy de acuerdo", "de acuerdo", "en desacuerdo", "muy en desacuerdo", "no sabe"}{"muy de acuerdo", "de acuerdo", "en desacuerdo", "muy en desacuerdo", "no sabe"}{"continuar como hasta ahora", "continuar con algunos cambios", "cambiar manteniendo solo algunas cosas", "cambiar totalmente", "no sabe"}{'muy informado', 'bastante informado', 'poco informado', 'nada informado', 'no sabe'}{'muy grave', 'bastante grave', 'poco grave', 'nada grave', 'no sabe'}
#Potencialidad:{"muy probable", "bastante probable", "poco probable", "nunca lo votaria", "no sabe"}
#Clase_social:{"Clase baja""Clase media baja""Clase media""Clase media alta""Clase alta""No sabe"}
#Principal_motivo_del_voto: "A favor de Milei""En contra de Milei""A favor del kirchnerismo""En contra del kirchnerismo ""A favor de la lista de Federico Achával""En contra de la lista de Federico Achával""No sabe"
#Afirmativa: ["si", "no", "no sabe"]
#nominal: "any other set of values"
