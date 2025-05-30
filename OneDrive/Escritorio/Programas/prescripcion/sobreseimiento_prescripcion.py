"""Utilidades para sobreseimientos por prescripción."""

TEMPLATE = """esta es la plantilla Córdoba, {fecha en letras}. 

VISTA: la presente causa caratulada {caratula}, venida a {este o esta} {tribunal} a los efectos de resolver la situación procesal de {nombre y apellido}, {datos personales} 

DE LA QUE RESULTA: Que {al/a la/a las/a los imputado/a/as/os} {nombre y apellido} se {le/s} atribuye {el siguiente hecho/los siguientes hechos}:  

{hecho} 

Y CONSIDERANDO: 

I. Que durante la instrucción se colectaron los siguientes elementos probatorios: 

{prueba} 

II. Que {el Sr. / la Sra.} {fiscal} requiere el sobreseimiento {total/parcial} en la presente causa respecto de {nombre y apellido}, por {el hecho/los hechos mencionado/s} supra, {encuadrado/s} bajo la calificación legal de {delitos}, en virtud de lo dispuesto por los arts. 348 y 350 inc. 4º del CPP, en función del art. 59 inc. 3º del CP, brindando los siguientes argumentos: {argumentos del fiscal} 

III. Conclusiones 

Analizada la cuestión traída a estudio, se advierte que {el hecho atribuido/los hechos atribuidos} a {nombre y apellido} {encuadra/n} efectivamente bajo la calificación legal de{delitos}, cuya pena máxima conminada en abstracto es de {penamaxima} de prisión. En este sentido, cabe aclarar que a los fines de computar el término para la prescripción del hecho imputado a {nombre y apellido} en los presentes autos se debe tener en cuenta {interrupción}, conforme surge de la planilla prontuarial, del Registro Nacional de Reincidencia y del Sistema de Administración de Causas. En efecto, {fundamentación} 

Así, teniendo en cuenta los términos referidos, entiendo que corresponde desvincular de la presente causa al imputado {nombre y apellido} por la causal de procedencia descripta en el art. 350 inc. 4º del CPP. Ello así, porque, tal como lo manifestó {el fiscal/la fiscal}, a la fecha, ha transcurrido con exceso el término establecido por el art. 62 inc. 2° del CP ({penamaxima} en este caso), el que desde la fecha {del hecho hasta la fecha del cumplimiento de la prescripción/de interrupción del plazo de la prescripción hasta su fecha del cumplimiento} no fue interrumpido por la comisión de nuevos delitos, conforme surge de la planilla prontuarial y del informe del Registro Nacional de Reincidencia incorporados digitalmente, y no procede ninguna de las causales contempladas por el art. 67 del CP, motivo por el cual ha de tenerse a la prescripción como causal de previo y especial pronunciamiento. Así lo establece el alto tribunal de esta provincia: “…Esta Sala, compartiendo la posición ya asumida por otra integración y por mayoría (A. nº 76, 29/6/93, "Cappa"; A. nº 60, 14/6/94, "Vivian"), ha sostenido que habida cuenta de la naturaleza sustancial de las distintas causales de sobreseimiento, las extintivas de la acción deben ser de previa consideración (T.S.J., Sala Penal, A. n° 26, 19/2/99, "Rivarola"; "Pérez", cit.). Por ello, la sola presencia de una causal extintiva de la acción -en el caso, la prescripción- debe ser estimada independientemente cualquiera sea la oportunidad de su producción y de su conocimiento por el Tribunal, toda vez que -en términos procesales- significa un impedimento para continuar ejerciendo los poderes de acción y de jurisdicción en procura de un pronunciamiento sobre el fondo (TSJ, Sala Penal, “CARUNCHIO, Oscar Rubén p.s.a. Homicidio Culposo -Recurso de Casación-” -Expte. "C", 36/03-, S. n.° 104 de fecha 16/9/2005). 

IV. En consecuencia, y de conformidad a lo normado por los arts. 59 inc. 3° y 62 inc. 2° del CP y 350 del CPP, corresponde declarar prescripta la pretensión punitiva penal emergente {del hecho calificado como configurativo/de los hechos calificados como configurativos} de {delitos} que se le {atribuía/n} a {nombre y apellido}.   

V. Finalmente, deberá oficiarse a la Policía de la Provincia de Córdoba y al Registro Nacional de Reincidencia a fin de informar lo aquí resuelto. 

Por lo expresado y disposiciones legales citadas; RESUELVO: 

I. Sobreseer {totalmente/parcialmente}, respecto {del hecho/de los hechos} de {fecha/s} {fechasdeloshechos}, a {nombre y apellido}, de condiciones personales ya relacionadas, por {el hecho calificado/los hechos calificados} como {delitos}, de conformidad con lo establecido por los arts. 348 y 350 inc. 4º del CPP, en función de los arts. 59 inc. 3º, 62 inc. 2º y 67 del CP. 

II. Ofícese a la Policía de la Provincia de Córdoba y al Registro Nacional de Reincidencia, a sus efectos. 

PROTOCOLÍCESE Y NOTIFÍQUESE. """


def render_prescripcion(**campos) -> str:
    """Devuelve el texto del sobreseimiento formateado."""
    return TEMPLATE.format(**campos)
