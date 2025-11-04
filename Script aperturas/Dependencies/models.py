
class question:
    def __init__(self, nombre, respuestas, bloque_df, ubicacion, tipo):
        self.nombre = None  # La pregunta
        self.respuestas = []  # Lista de respuestas (columna 1)
        self.bloque_df = None  # El bloque completo
        self.tipo = None  # Tipo de pregunta inferido
        self.ubicacion = None

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
    "nivel de responsabilidad": {'muy responsable', 'bastante responsable', 'poco responsable', 'nada responsable', 'no sabe'},
    "nivel de corrupccion": {'muy corrupto', 'bastante corrupto', 'poco corrupto', 'nada corrupto', 'no sabe'}
}

TITLE_CUES = {
    "imagen": ["imagen"],
    "gestion": ["como calificas", "gestion"],
    "multiple": ["multiple"],
    "clase social": ["clase social"],
    "estimulo": ["este video"],
    "afirmacion": ["afirmacion"],
    "principal motivo del voto": ["principal motivo"]
}

TEMPLATE_MAP = {
    "misc": 1,
    "imagen": 2,   # add missing labels you actually use
    "gestion": 2,
    "estimulo": 2,
    "valorativa_5": 2,
    "afirmacion": 3,
    "conocimiento": 3,
    "grado de informacion": 3,
    "nivel de gravedad": 3,
    "nivel de responsabilidad": 3,
    "nivel de corrupccion": 3,
    "valorativa_7": 4,
    "acuerdo": 5,
    "CoC": 5,
    "afirmativa": 6,
    "potencialidad": 7,
    "clase social": 8,
    "multiple": 9,
    "principal motivo del voto": 10
}