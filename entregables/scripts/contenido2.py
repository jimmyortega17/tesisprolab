# Parte 2: capítulos VI a VIII, ficha, apéndice y textos de actualización
import json
FIG = '../figuras/'
R = json.load(open('mc_resultados.json'))
def f0(x): return f'{x:,.0f}'
BASEPCT = round(100*sum(1 for x in R['van_array'] if x < 2061174.34)/len(R['van_array']))

def parte2(D):
    # ---------------- Capítulo VI: 6.1 ----------------
    D.pb(); D.h('Capítulo VI: Validación del Modelo de Negocio (secciones nuevas)', 1)
    D.where('Capítulo VI, dentro de 6.1 Validación de la Deseabilidad, después de la Tabla 15 y antes de 6.1.1.')
    D.at('c6a'); D.h('6.1.2 Mapa de Priorización de Hipótesis', 3)
    D.p('Las ocho hipótesis de la Tabla 15 no pesan igual. Se priorizaron con dos criterios: la importancia para el VAN y para la decisión de escalar, y la evidencia disponible. La puntuación es de 0 a 10 y es un juicio del equipo basado en el modelo financiero y en la matriz de supuestos críticos (Apéndice H). La Figura 6 muestra el mapa y la Tabla 58 explica cada puntaje.')
    D.figure('Mapa de priorización de hipótesis', FIG + 'fig_mapa_hipotesis.png', note='Elaboración propia. Las hipótesis del cuadrante superior izquierdo son importantes y tienen poca evidencia; se prueban primero.', width=5.8)
    D.table('Puntaje de priorización de las hipótesis', ['Hipótesis', 'Importancia', 'Evidencia', 'Fundamento', 'Orden de prueba'], [
        ['H5. Preventa de S/ 140', '10', '1', 'Sin ella no hay ingresos; el VAN se anula con S/ 104 y el precio de S/ 140 todavía no se ha probado con compra real.', '1'],
        ['H6. El productor firma', '9', '3', 'Sin café selecto no se cumple la preventa. Hay interés (85% acepta visitas), pero no firma.', '2'],
        ['H8. Recompra y visita', '8.5', '1.5', 'Define la retención de 60% y el LTV. Sin evidencia propia.', '3'],
        ['H4. Rendimiento de 22 qq/ha', '8', '3', 'Define cuántos productores se necesitan. Hay referencias externas, no medidas propias.', '4'],
        ['H7. Valora la trazabilidad', '6.5', '1', 'Sostiene la diferenciación; su efecto sobre el precio no está medido.', '5'],
        ['H1. Valora el pago al productor', '6', '6.5', '92% en el trabajo de campo (n = 25), con muestra por conveniencia.', 'Validada'],
        ['H3. El productor acepta visitas', '4', '6.5', '85% (n = 20). Es complemento del modelo.', 'Validada'],
        ['H2. Paga S/ 100 o más', '5', '5', '76% (n = 25), pero para una experiencia turística, no para el paquete de S/ 140.', 'Validar con H5'],
    ], note='Elaboración propia a partir de la Tabla 15 y del Apéndice H.', widths=[1.5, 0.8, 0.7, 2.7, 0.8])

    D.h('6.1.3 Diseño de los Experimentos', 3)
    D.p('Cada hipótesis tiene un experimento definido antes de conocer su resultado. La Tabla 59 detalla el tipo de experimento, la muestra, el canal, la duración, el costo y la decisión asociada. Los costos provienen del presupuesto del piloto (Tabla 32); donde no hay un costo específico se indica que está incluido en otro rubro.')
    D.table('Diseño de los experimentos por hipótesis', ['Hipótesis', 'Experimento', 'Muestra y canal', 'Duración', 'Costo (S/)', 'Criterio y decisión'], [
        ['H5', 'Preventa con pago real y prueba de precio (S/ 140 y S/ 130)', 'Unas 10,000 visitas calificadas desde pauta digital', 'Meses 2 a 4', '14,000 (página 8,000 y pauta 6,000)', 'Conversión de 3% y 300 paquetes. Si no se cumple, probar S/ 130 y luego S/ 120; no escalar si no se sostiene S/ 104.'],
        ['H6', 'Entrevista y propuesta de acuerdo escrito', '30 productores de Chanchamayo (Apéndice C)', 'Meses 1 a 2', 'Incluido en viajes (4,000)', '15 de 30 firman. Si no, revisar prima y condiciones y ampliar a Satipo.'],
        ['H8', 'Seguimiento de compradores del piloto', 'Compradores de los 300 paquetes', 'Meses 4 a 6', 'Incluido en el equipo', 'Recompra de 60% y 5% reserva tour. Si no, ajustar frecuencia y beneficios.'],
        ['H4', 'Fincas piloto con kit y asistencia técnica', '10 fincas; pesaje de la cosecha por parcela', 'Segunda campaña', '35,000 (kits) y 18,750 (técnico)', '22 qq/ha o más. Si no, recalcular cohortes y número de productores.'],
        ['H7', 'Encuesta de disposición a pagar y comparación con y sin videos', '150 compradores de café de especialidad (Apéndice B)', 'Meses 3 a 5', 'Incluido en pauta y equipo', '60% paga más por trazabilidad. Si no, simplificar videos y bajar su costo (3,000).'],
        ['H1 a H3', 'Ya realizados: encuesta de campo', '20 productores y 25 turistas', 'Abril a mayo', '—', 'Validadas; se contrastan con Lima en H5 y H7.'],
    ], note='Elaboración propia a partir de las Tablas 32 y 33 y de los Apéndices B y C.', widths=[0.6, 1.5, 1.3, 0.8, 1.0, 1.7], size=8.5)

    D.h('6.1.4 Tarjetas de Prueba y de Aprendizaje', 3)
    D.p('Cada hipótesis se resume en una tarjeta de prueba y una tarjeta de aprendizaje, según el formato de Bland y Osterwalder (2019). La tarjeta de prueba se completa antes del experimento y la de aprendizaje después. En H1 a H3 se registran los resultados de campo; en H4 a H8 los campos de resultado quedan en blanco hasta ejecutar el piloto.')
    cards = [
        ('H1', 'Valoración del pago al productor',
         'El consumidor valora que su pago beneficie directamente al productor.', 'Encuesta de campo a 25 turistas potenciales.', 'Porcentaje que responde que es importante.', '70% o más.',
         'Alta (92%, 23 de 25).', 'Tuvo la valoración más alta, por encima del interés en la experiencia (88%).', 'Usar el beneficio al productor como mensaje central de la página de preventa y confirmarlo con compradores de Lima.'),
        ('H2', 'Disposición a pagar',
         'El consumidor paga S/ 100 o más por una experiencia vinculada al café.', 'Encuesta de campo a 25 turistas potenciales.', 'Porcentaje que declara esa disposición.', '60% o más.',
         '76% (19 de 25).', 'La disposición es declarada y corresponde a una experiencia turística, no al paquete de S/ 140 de Lima.', 'Contrastar con la compra real de H5 antes de fijar el precio.'),
        ('H3', 'Aceptación de visitas',
         'El productor acepta recibir visitantes en su finca.', 'Encuesta de campo a 20 productores.', 'Porcentaje que responde afirmativamente.', '70% o más.',
         '85% (17 de 20).', 'El productor ve el turismo como medio para mejorar el valor de su café (80%).', 'Mantener las visitas como complemento y pactar el pago por persona en el acuerdo.'),
        ('H4', 'Rendimiento tecnificado',
         'La tecnificación eleva el rendimiento a 22 qq/ha o más.', 'Diez fincas piloto con kit y asistencia; pesaje de la cosecha por parcela.', 'Rendimiento por hectárea en la segunda campaña.', '22 qq/ha o más.',
         '[Completar tras la segunda campaña]', '[Completar]', '[Completar: si no se cumple, recalcular cohortes y el número de productores]'),
        ('H5', 'Preventa del paquete',
         'El consumidor de Lima compra por adelantado un paquete de S/ 140.', 'Página de preventa con pago real y pauta digital para unas 10,000 visitas calificadas.', 'Conversión de visitas calificadas y paquetes vendidos.', '3% y 300 paquetes.',
         '[Completar tras el piloto]', '[Completar]', '[Completar: probar S/ 130 y S/ 120, o detener según el precio de quiebre de S/ 104]'),
        ('H6', 'Compromiso del productor',
         'El productor firma el compromiso de venta de su café selecto.', 'Propuesta de acuerdo a 30 productores (Apéndice C).', 'Productores que firman.', '50% (15 de 30).',
         '[Completar tras el piloto]', '[Completar]', '[Completar: revisar prima y condiciones; ampliar a Satipo]'),
        ('H7', 'Valoración de la trazabilidad',
         'El consumidor valora la trazabilidad con videos y código QR.', 'Encuesta a 150 compradores de café de especialidad (Apéndice B) y comparación de la página con y sin videos.', 'Porcentaje que pagaría más por trazabilidad.', '60% o más.',
         '[Completar tras el piloto]', '[Completar]', '[Completar: simplificar videos y reducir su costo]'),
        ('H8', 'Recompra y visita',
         'El cliente recompra y visita la finca.', 'Seguimiento de los compradores del piloto durante seis meses.', 'Porcentaje de recompra y de reserva de tour.', '60% de recompra y 5% de reservas.',
         '[Completar tras el piloto]', '[Completar]', '[Completar: ajustar frecuencia y beneficios de recompra]'),
    ]
    for k, tit, creemos, prueba, metrica, crit, obs, aprend, decision in cards:
        D.card(f'Tarjeta de prueba y de aprendizaje de {k}: {tit}', [
            ('Tarjeta de prueba', ''), ('Creemos que', creemos), ('Para verificarlo haremos', prueba), ('Y mediremos', metrica), ('Estamos en lo correcto si', crit),
            ('Tarjeta de aprendizaje', ''), ('Observamos que', obs), ('Aprendimos que', aprend), ('Por lo tanto, decidimos', decision)],
            note='Elaboración propia. Formato adaptado de Bland y Osterwalder (2019).' if k == 'H1' else None)
    D.p('Referencia nueva para la lista: Bland, D. J., & Osterwalder, A. (2019). Testing business ideas: A field guide for rapid experimentation. John Wiley & Sons.', italic=True, size=10)

    D.h('6.1.5 Tareas y Preguntas de Prueba', 3)
    D.p('Antes de lanzar los experimentos se prueba con usuarios cada pieza del PMV de la segunda versión. La Tabla 68 resume las tareas y las preguntas que se aplican; el guion completo está en el Apéndice J.')
    D.table('Resumen de tareas de prueba y preguntas', ['Prueba', 'Hipótesis', 'Tarea que realiza el usuario', 'Pregunta de verificación'], [
        ['Página de preventa', 'H5, H7', 'Encontrar el precio y la fecha de entrega; elegir un productor; completar la compra.', '¿Qué información necesitaba para confiar en pagar por adelantado?'],
        ['Código QR y videos', 'H7', 'Escanear la bolsa y hallar la finca, la fecha de tueste y el video de la cosecha.', '¿Qué le pareció más útil y qué no entendió?'],
        ['Acuerdo con el productor', 'H6, H3', 'Leer el acuerdo y explicarlo con sus palabras; señalar qué condición cambiaría.', '¿Qué le haría desconfiar y firmaría hoy?'],
        ['Seguimiento posventa', 'H8', 'Recordar cuándo recibió su café y decir si renovaría su paquete.', '¿Qué necesitaría para volver a comprar o visitar la finca?'],
    ], note='Elaboración propia. Guion completo en el Apéndice J.', widths=[1.2, 0.8, 2.5, 2.0])

    # ---------------- 6.3.10 Monte Carlo ----------------
    D.pb(); D.where('Capítulo VI, al final de 6.3, después de 6.3.9 Análisis de Sensibilidad, Escenarios y Umbrales.')
    D.at('c6b'); D.h('6.3.10 Simulación Monte Carlo del VAN y del LTV', 3)
    D.p('El análisis de sensibilidad de la sección 6.3.9 cambia una variable por vez y usa cuatro escenarios con probabilidades asignadas por el equipo. La simulación Monte Carlo completa ese análisis: varía a la vez los supuestos de confianza media y baja del Apéndice H, extrae 10,000 combinaciones y calcula la distribución del VAN y del LTV. Se usaron distribuciones triangulares, con un mínimo, un valor más probable (el supuesto del modelo) y un máximo. Los rangos se apoyan en los escenarios del equipo y en los umbrales de las hipótesis; son supuestos que el piloto debe reemplazar por datos medidos. El COK se mantuvo fijo en 20.94%.')
    rows = []
    for k, lab, (lo, mo, hi), jus in R['dist']:
        fmt = (lambda v: f'{v*100:.1f}%') if (mo < 1 and k != 'paqf' and k != 'f_cf' and k != 'f_var') else (lambda v: f'{v:g}')
        rows.append([lab, fmt(lo), fmt(mo), fmt(hi), jus])
    D.table('Distribuciones de los supuestos simulados', ['Variable', 'Mínimo', 'Más probable', 'Máximo', 'Justificación'], rows,
            note='Elaboración propia. El costo de adquisición se simuló como costo por visita calificada (S/ 2.40) dividido entre la conversión simulada. Simulación con semilla fija en Python; modelo validado contra el Excel con el precio inicial de S/ 120 (VAN base de S/ 922,276 y los cuatro escenarios) y luego actualizado al precio de S/ 140.', widths=[1.9, 0.6, 0.7, 0.6, 2.7], size=8.5)
    v, l, c = R['van'], R['ltv'], R['ltv_cac']
    D.table('Resultados de la simulación Monte Carlo (10,000 iteraciones)', ['Estadístico', 'VAN (S/)', 'LTV (S/ por cliente)', 'LTV / CAC'], [
        ['Media', f0(v['media']), f0(l['media']), f"{c['media']:.1f}"],
        ['Mediana', f0(v['mediana']), f0(l['mediana']), f"{c['mediana']:.1f}"],
        ['Desviación estándar', f0(v['desv']), f0(l['desv']), f"{c['desv']:.1f}"],
        ['Percentil 5', f0(v['p5']), f0(l['p5']), f"{c['p5']:.1f}"],
        ['Percentil 25', f0(v['p25']), f0(l['p25']), f"{c['p25']:.1f}"],
        ['Percentil 75', f0(v['p75']), f0(l['p75']), f"{c['p75']:.1f}"],
        ['Percentil 95', f0(v['p95']), f0(l['p95']), f"{c['p95']:.1f}"],
        ['Probabilidad de VAN negativo', f"{R['p_van_neg']*100:.1f}%", '—', '—'],
        ['Probabilidad de LTV/CAC menor que 3', '—', '—', f"{R['p_lc_menor3']*100:.1f}%"],
    ], note='Elaboración propia. Caso base del modelo: VAN de S/ 2,061,174 y LTV de S/ 793. Resultados en la hoja MonteCarlo del Excel.', widths=[2.6, 1.3, 1.5, 1.1])
    D.figure('Histograma del VAN simulado', FIG + 'fig_mc_van.png', note='Elaboración propia. Las barras en rojo corresponden a VAN negativo.', width=5.8)
    D.figure('Histograma del LTV simulado', FIG + 'fig_mc_ltv.png', note='Elaboración propia.', width=5.8)
    D.figure('Histograma de la relación LTV/CAC simulada', FIG + 'fig_mc_ltvcac.png', note='Elaboración propia. Las barras en rojo corresponden a una relación menor que 3.', width=5.8)
    D.p(f"Resultados. El VAN simulado tiene mediana de S/ {f0(v['mediana'])} y media de S/ {f0(v['media'])}, inferiores al caso base de S/ 2,061,174, que queda por encima de la mediana de la simulación (percentil {BASEPCT}). La probabilidad de que el VAN sea negativo es de {R['p_van_neg']*100:.1f}%. En el percentil 5 el VAN sigue siendo positivo, de S/ {f0(v['p5'])}. El VAN esperado por escenarios de S/ 1,713,197 de la Tabla 28 resulta más optimista que la media simulada; se recomienda reportar ambos y citar la simulación como el resultado más prudente. Estos resultados dependen de que el consumidor pague S/ 140, lo que H5 debe confirmar.")
    D.p(f"El LTV por cliente tiene media de S/ {f0(l['media'])} y un rango de S/ {f0(l['p5'])} a S/ {f0(l['p95'])} entre los percentiles 5 y 95, por debajo del valor base de S/ 793. La relación LTV/CAC tiene mediana de {c['mediana']:.1f} y solo {R['p_lc_menor3']*100:.1f}% de las iteraciones cae por debajo de 3, por lo que la economía del cliente es más resistente que el VAN total: el riesgo del proyecto se concentra en que el consumidor acepte el precio y en la escala, no en el valor de cada cliente.")
    D.figure('Variables que más explican la variación del VAN', FIG + 'fig_mc_correlacion.png', note='Elaboración propia. Correlación de rangos de Spearman entre cada variable simulada y el VAN.', width=5.8)
    D.p(f"La Figura 10 confirma que el precio del paquete es la variable que más explica el VAN (correlación de {R['corr_van']['precio']:.2f}), seguido del precio pagado por el café (correlación de {R['corr_van']['p_emp']:.2f}), la conversión de visitas ({R['corr_van']['conv']:.2f}) y el volumen vendido ({R['corr_van']['paqf']:.2f}). Esto coincide con el tornado de la Figura 11 y con el precio de quiebre de S/ 104, y explica por qué H5 (preventa de S/ 140) es la hipótesis que debe probarse primero: el VAN sigue siendo positivo con precios de hasta S/ 105, pero solo si el consumidor acepta pagar más que la referencia de mercado.")
    D.figure('Tornado de sensibilidad del VAN', FIG + 'fig_tornado.png', note='Elaboración propia a partir del modelo financiero. Cada variable se mueve ±20% con las demás en su valor base.', width=5.8)
    D.p('Limitaciones. Las variables se simularon de forma independiente, aunque en la práctica algunas se mueven juntas (por ejemplo, un precio menor suele acompañarse de menor volumen), lo que podría ampliar el riesgo. Los rangos son juicios del equipo y deben actualizarse con los resultados del piloto. La simulación no modela el riesgo de un mal año climático ni el COK variable.')

    D.h('6.3.11 Recupero, TIR Modificada y Plan de Financiamiento', 3)
    D.p('La Tabla 71 completa los indicadores de la Tabla 25 con el recupero simple, la TIR modificada y la necesidad de financiamiento. Se calculan en la hoja Indicadores_Extra del Excel.')
    D.table('Indicadores adicionales de rentabilidad y financiamiento', ['Indicador', 'Valor'], [
        ['Periodo de recupero simple', '1.48 años'], ['Periodo de recupero descontado', '1.74 años'],
        ['TIR', '128.8%'], ['TIR modificada (financiamiento y reinversión al COK de 20.94%)', '85.9%'], ['Índice de rentabilidad', '1.20'], ['Flujo de caja libre del año 1', 'S/ 48,543'],
        ['Aporte inicial de los socios (año 0)', 'S/ 271,700'],
        ['Necesidad máxima acumulada de caja (año 0)', 'S/ 271,700'],
        ['Aporte total recomendado con 20% de colchón', 'S/ 326,040'],
        ['Aporte por socio con cuatro socios', 'S/ 81,510'],
    ], note='Elaboración propia a partir de la hoja FCL. El colchón de 20% es una propuesta del equipo.', widths=[4.5, 2.0])
    D.p('La TIR modificada de 85.9% es menor que la TIR de 128.8% porque supone reinvertir los flujos al COK y no a la propia TIR; es una medida más prudente. Con el precio de S/ 140, el flujo del año 1 es positivo (S/ 48,543) y la necesidad máxima de caja coincide con el aporte inicial del año 0, de S/ 271,700. El anticipo de los clientes por la preventa cubre el inventario, por lo que el modelo no requiere capital de trabajo adicional. El plan de financiamiento propone cubrir la inversión con aporte de los socios y agregar un colchón de 20%.')
    D.table('Plan de financiamiento por etapas', ['Etapa', 'Uso', 'Monto (S/)', 'Fuente'], [
        ['0. Piloto', 'Presupuesto del piloto (Tabla 32), menos S/ 30,500 de ventas de los 300 paquetes de prueba', '63,050 netos (93,550 brutos)', 'Aporte de los socios'],
        ['1. Lanzamiento (año 0)', 'Inversión inicial y kits del año 1', '271,700', 'Aporte de los socios, solo si se decide escalar'],
        ['2. Colchón de seguridad', 'Reserva de 20% sobre la necesidad máxima, por si la preventa se retrasa', '54,340', 'Aporte de los socios'],
    ], note='Elaboración propia. El modelo no descuenta del año 0 los kits y el contenido del piloto; si se reutilizan, la necesidad de caja baja. El total con colchón (S/ 326,040) es el aporte a comprometer antes de escalar. El proyecto no usa deuda ni fondos del Estado.', widths=[1.3, 2.5, 1.5, 1.2])

    # ---------------- Capítulo VII ----------------
    D.pb(); D.h('Capítulo VII: Solución Sostenible (secciones nuevas)', 1)
    D.where('Capítulo VII, después de 7.2 Rentabilidad Social de la Solución.')
    D.at('c7'); D.h('7.3 Impacto Ambiental', 2)
    D.p('El modelo no reduce su sostenibilidad a la dimensión social. La Tabla 73 identifica los impactos ambientales de cada etapa y las medidas de gestión. Las metas son propuestas del equipo y se miden en las fincas piloto.')
    D.table('Impactos ambientales y medidas de gestión', ['Etapa', 'Impacto', 'Medida', 'Indicador'], [
        ['Producción', 'Uso de agua y fertilizantes; control de plagas.', 'Fertirriego por goteo con dosis medidas; manejo integrado de plagas; sombra y cobertura del suelo.', 'Litros de agua y kilos de fertilizante por quintal'],
        ['Beneficio y secado', 'Pulpa, aguas de lavado y secado con leña o gas.', 'Secado solar; compostaje de la pulpa; manejo de las aguas de lavado.', 'Porcentaje de lotes con secado solar y de pulpa compostada'],
        ['Tueste y empaque', 'Consumo de energía; bolsas con válvula y cajas.', 'Tueste por lotes con aliado; empaque reciclable cuando sea posible; cajas del tamaño exacto.', 'Kilos de empaque por paquete'],
        ['Transporte y envío', 'Emisiones del flete de Chanchamayo a Lima y del envío a domicilio.', 'Consolidar cargas y despachos por rutas; elegir mensajería con prácticas ambientales.', 'Envíos por ruta consolidada'],
        ['Cliente', 'Residuos de la caja y de las bolsas.', 'Indicaciones de reciclaje en el código QR.', 'Clientes que reciben las indicaciones'],
    ], note='Elaboración propia. Pendiente de medir en las fincas piloto.', widths=[1.1, 1.7, 2.4, 1.3])
    D.h('7.4 Enfoque de Género e Inclusión', 2)
    D.p('La caficultura familiar involucra a mujeres y jóvenes en la producción, el beneficio y la venta. El modelo incorpora el enfoque de género como compromiso de gestión, con indicadores que se registran desde el piloto. La Tabla 74 resume las medidas y las metas propuestas por el equipo.')
    D.table('Medidas de género e inclusión', ['Medida', 'Indicador', 'Meta propuesta'], [
        ['Invitar a las productoras y a los hijos jóvenes a las capacitaciones técnicas y fijar horarios compatibles con el trabajo familiar.', 'Mujeres y jóvenes capacitados', 'Al menos 30% de los participantes del piloto'],
        ['Pagar el café y las visitas a quien produce, en una cuenta a nombre del titular o de la pareja, según el acuerdo.', 'Pagos a mujeres titulares', 'Registrar el 100% de los pagos'],
        ['Mostrar en los videos y las fichas a las mujeres y jóvenes que participan, con su consentimiento.', 'Fichas con mujeres o jóvenes visibles', 'Al menos 30% de las fichas'],
        ['Contratar localmente para los puestos de campo y de almacén.', 'Puestos ocupados por mujeres', 'Revisar cada año'],
    ], note='Elaboración propia. Las metas son propuestas del equipo y se confirman con los datos de los productores del piloto.', widths=[3.3, 1.7, 1.5])

    # ---------------- Capítulo VIII ----------------
    D.pb(); D.h('Capítulo VIII: Decisión e Implementación (sección nueva)', 1)
    D.where('Capítulo VIII, dentro de 8.1 Plan de Implementación y Equipo de Trabajo, después de la Tabla 33.')
    D.at('c8'); D.h('8.1.1 Estructura Organizacional', 3)
    D.p('La Figura 12 presenta la estructura organizacional del primer año. Los cuatro fundadores forman la junta de socios y asumen las áreas de gerencia general, gestión comercial y digital, finanzas y relación con los aliados. Los puestos de campo, ventas, almacén y comunicación se contratan según el modelo de costos fijos. La Tabla 75 resume los roles.')
    D.figure('Organigrama de Cosecha Viva', FIG + 'fig_organigrama.png', note='Elaboración propia a partir de la hoja CostosFijos del modelo financiero.', width=6.0)
    D.table('Roles y costos del equipo', ['Puesto', 'Función principal', 'Costo mensual (S/)'], [
        ['Gerente general', 'Estrategia, alianzas, relación con productores y socios', '4,500'],
        ['Jefe agronómico', 'Asistencia técnica, calidad y medición de rendimientos (H4)', '4,000'],
        ['Comunicador audiovisual', 'Videos por etapa, fichas y contenido', '3,000'],
        ['Ejecutivo de ventas y atención', 'Preventa, clientes y regalos corporativos (H5, H8)', '2,200'],
        ['Asistente de almacén', 'Recepción, empaque y despachos', '1,500'],
        ['Técnico de campo', 'Un técnico por cada 30 productores (2 en el año 1; 4 en el año 5)', '2,500 cada uno'],
    ], note='Adaptado de la hoja CostosFijos del modelo. Los sueldos son un supuesto del equipo y se suman cargas laborales de 25%.', widths=[1.8, 3.4, 1.3])

    # ---------------- Ficha, apéndice y textos ----------------
    D.at('END'); D.h('Apéndice I. Ficha Resumen del Proyecto', 1)
    D.where('Después del Resumen Ejecutivo o como apéndice, como hoja de una página.')
    D.table('Ficha resumen de Cosecha Viva', ['Elemento', 'Contenido'], [
        ['Problema', 'Los pequeños productores de café de Chanchamayo obtienen ingresos bajos e inestables: producen cerca de 15 qq/ha, venden sin diferenciar a intermediarios y no llegan al consumidor urbano.'],
        ['Solución', 'Tecnificación con kit en comodato y asistencia técnica; compra del café selecto a S/ 20 por kg de pergamino (33% sobre el intermediario); venta por preventa a consumidores de Lima y la costa.'],
        ['Producto', 'Paquete de cosecha de 1 kg a S/ 140 con videos por etapa, código QR por lote y visita opcional a la finca.'],
        ['Clientes', 'Consumidor de café de especialidad de Lima y la costa (351,000 personas interesadas); regalos corporativos; visitantes.'],
        ['Ingresos', 'Paquetes de cosecha y tours a la finca (S/ 180 por persona).'],
        ['Inversión', 'S/ 156,200 de inversión inicial; aporte inicial de S/ 271,700; necesidad máxima de caja de S/ 271,700 (flujo del año 1 de +S/ 48,543).'],
        ['Rentabilidad (base)', 'VAN de S/ 2,061,174; TIR de 128.8%; recupero simple de 1.5 años; COK de 20.94%.'],
        ['Riesgo (Monte Carlo)', f"VAN mediano de S/ {f0(R['van']['mediana'])}; probabilidad de VAN negativo de {R['p_van_neg']*100:.1f}%; el precio sigue siendo la variable crítica (quiebre en S/ 104)."],
        ['Impacto social', 'Ingreso neto del productor +64.7%; VAN social al 8% de S/ 7,370,976.'],
        ['Evidencia', 'H1 a H3 validadas en campo; H4 a H8 pendientes de un piloto de diez fincas y 300 paquetes de preventa.'],
        ['Decisión', 'Ejecutar el piloto de seis meses (S/ 93,550) y escalar solo si H5 y H6 se cumplen y el VAN recalculado es positivo.'],
    ], note='Elaboración propia a partir de la tesis y del modelo financiero v5.', widths=[1.4, 5.1])

    D.h('Apéndice J. Guion de las Tareas de Prueba', 1)
    D.where('Después del Apéndice H. Actualice el índice de apéndices.')
    D.p('Cada prueba se aplica a entre 5 y 8 personas antes del experimento principal, con un moderador y un observador. Se registra lo que hace la persona, lo que dice y dónde se detiene. No se explica la pantalla; se pide pensar en voz alta. Al final se piden permiso para usar sus respuestas y la grabación.')
    def guion(titulo, hip, perfil, tareas, preguntas, exito):
        D.h(titulo, 3)
        D.p(f'Hipótesis: {hip}. Participantes: {perfil}.', size=11)
        D.p('Tareas', bold=True, size=11); D.bullets(tareas)
        D.p('Preguntas después de cada tarea y al final', bold=True, size=11); D.bullets(preguntas)
        D.p(f'Señal de éxito: {exito}', italic=True, size=11)
    guion('I.1 Prueba de la página de preventa', 'H5 y H7', 'consumidores de café de 25 a 55 años que compran café al menos una vez al mes (5 a 8)',
          ['Encuentre cuánto cuesta el paquete y cuándo se entrega.', 'Elija el productor cuyo café quiere recibir.', 'Complete la compra hasta el pago, sin pagar.', 'Diga cómo pediría la devolución si algo sale mal.'],
          ['¿Qué esperaba ver antes de pagar y no encontró?', '¿Qué le generó dudas sobre pagar por adelantado?', 'Del 1 al 5, ¿qué tan seguro se siente de que recibirá su café?', '¿Pagaría S/ 140? Si dice que no, ¿qué precio le parecería justo?', '¿Qué le hizo falta para decidir?'],
          'Al menos 70% completa la compra sin ayuda y el promedio de seguridad es de 4 o más.')
    guion('I.2 Prueba del código QR y de los videos', 'H7', 'los mismos consumidores y 3 baristas o compradores frecuentes',
          ['Escanee el código de la bolsa de muestra.', 'Encuentre la finca, la fecha de tueste y el nombre del productor.', 'Mire un video de una etapa y diga qué entendió.'],
          ['¿Qué le pareció más útil de lo que vio?', '¿Qué no entendió o le sobró?', '¿Esta información cambia lo que pagaría por el café? ¿Cuánto más?', '¿Con qué frecuencia vería estos videos?'],
          'El 60% o más dice que pagaría más y encuentra la información sin ayuda.')
    guion('I.3 Prueba del acuerdo con productores', 'H6, H3 y H4', 'productores de Chanchamayo (10 a 15 en la primera ronda)',
          ['Lea el acuerdo y cuéntelo con sus palabras.', 'Señale qué condición cambiaría.', 'Diga qué haría si el mercado paga más que el precio pactado.', 'Describa cómo recibiría la asistencia y el kit.'],
          ['¿Qué le daría confianza para firmar?', '¿Qué le haría vender a otro comprador?', '¿Qué necesitaría para producir más café selecto?', '¿Acepta que grabemos videos de su finca? ¿Y recibir visitantes?'],
          'El 50% o más firma; las condiciones que piden cambiar se resumen en tres o menos.')
    guion('I.4 Seguimiento de compradores', 'H8', 'compradores del piloto (10 entrevistas a los 2 y a los 5 meses)',
          ['Recuerde cuándo recibió su café y cómo fue la entrega.', 'Diga si renovaría su paquete y con qué frecuencia.', 'Diga si reservaría una visita a la finca.'],
          ['¿Cómo preparó el café y qué le pareció?', '¿Qué lo haría comprar de nuevo?', '¿Compartió los videos con alguien? ¿Regalaría un paquete?', '¿Qué cambiaría del servicio?'],
          '60% recompra y 5% reserva tour; satisfacción de 4 o más.')

