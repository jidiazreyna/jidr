"""Utilidades para sobreseimientos por prescripción."""

from collections import OrderedDict


# Campos que componen la plantilla.  La clave es el nombre que debe
# pasarse a ``render_prescripcion`` y el valor es una breve descripción
# mostrable en la interfaz.
CAMPOS = OrderedDict([
    ("fecha_letras", "fecha en letras"),
    ("caratula", "carátula"),
    ("este_esta", "este o esta"),
    ("tribunal", "tribunal"),
    ("nombre_apellido", "nombre y apellido"),
    ("datos_personales", "datos personales"),
    ("imputado_articulo", "al/a la/a las/a los imputado/a/as/os"),
    ("le_les", "le/s"),
    ("hechos_label", "el siguiente hecho/los siguientes hechos"),
    ("hecho", "hecho"),
    ("prueba", "prueba"),
    ("titulo_fiscal", "el Sr. / la Sra."),
    ("fiscal", "fiscal"),
    ("alcance", "total/parcial"),
    ("hechos_mencionados", "el hecho/los hechos mencionado/s"),
    ("encuadrado", "encuadrado/s"),
    ("delitos", "delitos"),
    ("argumentos_fiscal", "argumentos del fiscal"),
    ("hechos_atribuidos", "el hecho atribuido/los hechos atribuidos"),
    ("encuadra", "encuadra/n"),
    ("penamaxima", "penamaxima"),
    ("interrupcion", "interrupción"),
    ("fundamentacion", "fundamentación"),
    ("el_fiscal", "el fiscal/la fiscal"),
    (
        "periodo_prescripcion",
        "del hecho hasta la fecha del cumplimiento de la prescripción/"
        "de interrupción del plazo de la prescripción hasta su fecha del cumplimiento",
    ),
    ("hecho_calif", "del hecho calificado como configurativo/" "de los hechos calificados como configurativos"),
    ("atribuia", "atribuía/n"),
    ("fechas_hechos", "fechasdeloshechos"),
    ("sobreseimiento_total_parcial", "totalmente/parcialmente"),
    ("hecho_hechos", "del hecho/de los hechos"),
    ("fecha_s", "fecha/s"),
    ("hecho_calificado", "el hecho calificado/los hechos calificados"),
])


# Plantilla de resolución por prescripción.  Se utilizan los nombres de
# ``CAMPOS`` como llaves para ``str.format``.
TEMPLATE = """
Córdoba, {fecha_letras}.

VISTA: la presente causa caratulada {caratula}, venida a {este_esta} {tribunal} a los efectos de resolver la situación procesal de {nombre_apellido}, {datos_personales}

DE LA QUE RESULTA: Que {imputado_articulo} {nombre_apellido} se {le_les} atribuye {hechos_label}:

{hecho}

Y CONSIDERANDO:

I. Que durante la instrucción se colectaron los siguientes elementos probatorios:

{prueba}

II. Que {titulo_fiscal} {fiscal} requiere el sobreseimiento {alcance} en la presente causa respecto de {nombre_apellido}, por {hechos_mencionados} supra, {encuadrado} bajo la calificación legal de {delitos}, en virtud de lo dispuesto por los arts. 348 y 350 inc. 4º del CPP, en función del art. 59 inc. 3º del CP, brindando los siguientes argumentos: {argumentos_fiscal}

III. Conclusiones

Analizada la cuestión traída a estudio, se advierte que {hechos_atribuidos} a {nombre_apellido} {encuadra} efectivamente bajo la calificación legal de{delitos}, cuya pena máxima conminada en abstracto es de {penamaxima} de prisión. En este sentido, cabe aclarar que a los fines de computar el término para la prescripción del hecho imputado a {nombre_apellido} en los presentes autos se debe tener en cuenta {interrupcion}, conforme surge de la planilla prontuarial, del Registro Nacional de Reincidencia y del Sistema de Administración de Causas. En efecto, {fundamentacion}

Así, teniendo en cuenta los términos referidos, entiendo que corresponde desvincular de la presente causa al imputado {nombre_apellido} por la causal de procedencia descripta en el art. 350 inc. 4º del CPP. Ello así, porque, tal como lo manifestó {el_fiscal}, a la fecha, ha transcurrido con exceso el término establecido por el art. 62 inc. 2° del CP ({penamaxima} en este caso), el que desde la fecha {periodo_prescripcion} no fue interrumpido por la comisión de nuevos delitos, conforme surge de la planilla prontuarial y del informe del Registro Nacional de Reincidencia incorporados digitalmente, y no procede ninguna de las causales contempladas por el art. 67 del CP, motivo por el cual ha de tenerse a la prescripción como causal de previo y especial pronunciamiento. Así lo establece el alto tribunal de esta provincia: “…Esta Sala, compartiendo la posición ya asumida por otra integración y por mayoría (A. nº 76, 29/6/93, "Cappa"; A. nº 60, 14/6/94, "Vivian"), ha sostenido que habida cuenta de la naturaleza sustancial de las distintas causales de sobreseimiento, las extintivas de la acción deben ser de previa consideración (T.S.J., Sala Penal, A. n° 26, 19/2/99, "Rivarola"; "Pérez", cit.). Por ello, la sola presencia de una causal extintiva de la acción -en el caso, la prescripción- debe ser estimada independientemente cualquiera sea la oportunidad de su producción y de su conocimiento por el Tribunal, toda vez que -en términos procesales- significa un impedimento para continuar ejerciendo los poderes de acción y de jurisdicción en procura de un pronunciamiento sobre el fondo (TSJ, Sala Penal, “CARUNCHIO, Oscar Rubén p.s.a. Homicidio Culposo -Recurso de Casación-” -Expte. "C", 36/03-, S. n.° 104 de fecha 16/9/2005).

IV. En consecuencia, y de conformidad a lo normado por los arts. 59 inc. 3° y 62 inc. 2° del CP y 350 del CPP, corresponde declarar prescripta la pretensión punitiva penal emergente {hecho_calif} de {delitos} que se le {atribuia} a {nombre_apellido}.

V. Finalmente, deberá oficiarse a la Policía de la Provincia de Córdoba y al Registro Nacional de Reincidencia a fin de informar lo aquí resuelto.

Por lo expresado y disposiciones legales citadas; RESUELVO:

I. Sobreseer {sobreseimiento_total_parcial}, respecto {hecho_hechos} de {fecha_s} {fechas_hechos}, a {nombre_apellido}, de condiciones personales ya relacionadas, por {hecho_calificado} como {delitos}, de conformidad con lo establecido por los arts. 348 y 350 inc. 4º del CPP, en función de los arts. 59 inc. 3º, 62 inc. 2º y 67 del CP.

II. Ofícese a la Policía de la Provincia de Córdoba y al Registro Nacional de Reincidencia, a sus efectos.

PROTOCOLÍCESE Y NOTIFÍQUESE.
"""


def render_prescripcion(**campos) -> str:
    """Devuelve el texto del sobreseimiento formateado."""
    return TEMPLATE.format(**campos)
