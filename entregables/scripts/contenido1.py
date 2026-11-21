# Parte 1: capítulos II a V
FIG = '../figuras/'

def parte1(D):
    D.h('Contenido nuevo para la tesis ProLab Cosecha Viva', 1)
    D.p('Este documento reúne las secciones que faltaban en la versión 4.1 de la tesis. Cada sección indica dónde insertarla. '
        'La numeración de tablas continúa desde la Tabla 42 y la de figuras desde la Figura 6; al pegar en la tesis, renumere según el orden final y actualice el índice y las listas de tablas y figuras. '
        'Las cifras financieras provienen del modelo Modelo_Financiero_Tesis_Cosecha_Viva_v5.xlsx. Las frases marcadas como "propuesta del equipo" son supuestos que el equipo debe confirmar o ajustar.')
    D.table('Mapa de inserción de las secciones nuevas', ['Sección nueva', 'Dónde va en la tesis'], [
        ['2.3 Análisis del entorno (PESTEL), 2.4 Cinco fuerzas, 2.5 Segmentación', 'Capítulo II, después de 2.2'],
        ['3.1.1 Mapas de empatía, 3.2.1 Mapa de experiencia del consumidor', 'Capítulo III, después de 3.1 y 3.2'],
        ['4.4.1 Lienzo de la propuesta de valor', 'Capítulo IV, dentro de 4.4'],
        ['5.5 Plan de marketing, 5.6 Aspectos legales, 5.3.1 Hoja de ruta de escalamiento y salida', 'Capítulo V'],
        ['6.1.2 a 6.1.5 Priorización, tarjetas de prueba y aprendizaje, experimentos y tareas de prueba', 'Capítulo VI, dentro de 6.1'],
        ['6.3.10 Simulación Monte Carlo, 6.3.11 Recupero, TIRM y financiamiento', 'Capítulo VI, al final de 6.3'],
        ['7.3 Impacto ambiental, 7.4 Género e inclusión', 'Capítulo VII'],
        ['8.1.1 Estructura organizacional', 'Capítulo VIII, dentro de 8.1'],
        ['Ficha resumen, Apéndice I (guiones de prueba), textos para el resumen y las conclusiones', 'Después del resumen; apéndices; capítulo VIII'],
    ], widths=[4.0, 2.5])

    # ---------------- Capítulo II ----------------
    D.pb(); D.h('Capítulo II: Análisis del Mercado (secciones nuevas)', 1)
    D.where('Capítulo II, después de la sección 2.2 Análisis Competitivo Detallado.')
    D.at('c2'); D.h('2.3 Análisis del Entorno', 2)
    D.p('El entorno se analizó con el modelo PESTEL, que ordena los factores políticos, económicos, sociales, tecnológicos, ecológicos y legales según su efecto sobre Cosecha Viva. La Tabla 43 presenta el resultado y la respuesta del modelo a cada factor.')
    D.table('Análisis PESTEL de Cosecha Viva', ['Factor', 'Situación', 'Efecto sobre el modelo', 'Respuesta de Cosecha Viva'], [
        ['Político', 'El riesgo país bajó a 118 puntos básicos (Andina, 2026). El modelo no usa fondos ni programas del Estado.', 'Bajo: no depende de subsidios ni de licitaciones.', 'Se financia con aporte de los socios y se mantiene el riesgo país dentro del COK.'],
        ['Económico', 'Precios históricos del café (Agraria.pe, 2026); inflación esperada de 2.25% (BCRP, 2026); consumo interno de 1.7 kg por persona (Agraria.pe, 2025).', 'Alto: un alza del precio en chacra reduce la prima de 33% y el margen de S/ 55.24 por paquete.', 'Precio pactado con prima sobre el precio de chacra del mes y piso garantizado; simulación del precio del café en el análisis de riesgo.'],
        ['Social', 'El consumidor valora el origen y que su pago beneficie al productor (92% en el trabajo de campo). El productor vende sin diferenciar a intermediarios.', 'Favorable: la propuesta responde a un valor de los dos usuarios.', 'Videos por etapa, código QR por lote y compromiso de compra con prima.'],
        ['Tecnológico', 'Pagos digitales, comercio electrónico, video y códigos QR son de uso común; el fertirriego y el secado solar son prácticas conocidas.', 'Favorable: la tecnología necesaria ya existe y es de bajo costo.', 'Plataforma de preventa y trazabilidad (S/ 45,000 con imprevistos S/ 49,500); kit de tecnificación en comodato.'],
        ['Ecológico', 'Roya, broca y lluvia en la época de secado afectan el volumen y la calidad (Tabla 8).', 'Medio: un mal año reduce el café selecto disponible.', 'Fincas en distintas altitudes, manejo integrado de plagas, secado solar y red de productores de respaldo.'],
        ['Legal', 'Venta a distancia con preventa, tratamiento de datos de clientes, comprobantes electrónicos y autorizaciones sanitarias del tueste.', 'Medio: el incumplimiento genera sanciones y daña la confianza.', 'Tostaduría con autorizaciones propias, contratos escritos y política de privacidad (sección 5.6).'],
    ], note='Elaboración propia a partir del modelo financiero, del trabajo de campo y de las fuentes citadas.', widths=[0.8, 2.0, 1.7, 2.0])
    D.h('2.4 Análisis de las Cinco Fuerzas Competitivas', 2)
    D.p('La sección 2.2 ya incluye un párrafo breve sobre las cinco fuerzas; esta sección lo amplía y puede reemplazarlo. Las fuerzas se evaluaron para el negocio de café de especialidad con preventa en Lima. La Tabla 44 resume la intensidad de cada fuerza y la respuesta del modelo. La valoración corresponde al juicio del equipo y debe contrastarse con entrevistas a tostadurías y cafeterías.')
    D.table('Cinco fuerzas competitivas', ['Fuerza', 'Intensidad', 'Fundamento', 'Respuesta'], [
        ['Rivalidad entre competidores', 'Alta', 'Hay más de 400 marcas de café de especialidad en Lima y modelos de venta anticipada como CrowdFarming (Tabla 4).', 'Diferenciarse por la combinación de tecnificación, preventa, videos y tours, no por precio.'],
        ['Amenaza de nuevos entrantes', 'Media', 'Abrir una tienda en línea es barato; replicar la relación con los productores y la tecnificación toma tiempo.', 'Acuerdos con productores, base de datos de clientes y marca ligada a las fincas.'],
        ['Poder de negociación de los proveedores', 'Alta', 'El productor puede vender a otro comprador cuando sube el precio en chacra.', 'Prima, piso garantizado, kit en comodato, asistencia técnica y red de respaldo.'],
        ['Poder de negociación de los clientes', 'Media', 'El consumidor compara con otras marcas y puede dejar de recomprar; la retención de 60% es un supuesto sin evidencia propia.', 'Videos, QR, regalos y beneficios de recompra; medir la retención en el piloto.'],
        ['Amenaza de sustitutos', 'Media', 'Café soluble, café de supermercado y cápsulas cubren la necesidad básica de café a menor precio.', 'Posicionar el paquete como experiencia de origen y regalo, no como café diario.'],
    ], note='Elaboración propia. Intensidad: valoración cualitativa del equipo.', widths=[1.4, 0.8, 2.3, 2.0])
    D.h('2.5 Segmentación y Mercado Objetivo', 2)
    D.p('El modelo atiende a cuatro segmentos. La Tabla 45 los describe y señala el criterio de prioridad. El segmento principal es el consumidor de café de especialidad de Lima y la costa, que es el mercado disponible de 351,000 personas de la Tabla 3. Los regalos corporativos y los visitantes de las fincas complementan la demanda; los productores son el segmento de oferta.')
    D.table('Segmentos del modelo Cosecha Viva', ['Segmento', 'Descripción', 'Necesidad principal', 'Prioridad y canal'], [
        ['Consumidor de café de especialidad', 'Personas de 25 a 55 años de Lima y la costa que compran café en grano o molido al menos una vez al mes (Apéndice B).', 'Café fresco y trazable, con origen visible e impacto en el productor.', 'Principal. Página de preventa, contenido en redes y referidos.'],
        ['Regalo corporativo y personal', 'Empresas y personas que regalan en fechas especiales.', 'Un regalo con historia y proveedor ético.', 'Complementario. Ventas directas B2B y campañas por fechas.'],
        ['Visitante de la finca', 'Clientes activos y sus acompañantes (2.5 personas por visita; 5% de los clientes).', 'Conocer la finca y al productor.', 'Segundo ingreso. Reserva desde la página y alianzas con hoteles y operadores.'],
        ['Productor de café de Chanchamayo', 'Productores con alrededor de dos hectáreas, rendimiento de 15 qq/ha y venta a intermediarios.', 'Mayor rendimiento, precio con prima y comprador seguro.', 'Oferta. Acuerdos escritos y asistencia técnica; 33 productores en el año 1 y 117 en el año 5.'],
    ], note='Elaboración propia a partir de las Tablas 3 y 5 y de los Apéndices B y C.', widths=[1.3, 2.0, 1.6, 1.6])

    # ---------------- Capítulo III ----------------
    D.pb(); D.h('Capítulo III: Investigación del Usuario (secciones nuevas)', 1)
    D.where('Capítulo III: el mapa de empatía va al final de 3.1 Perfil del Usuario; el mapa de experiencia del consumidor va al final de 3.2.')
    D.at('c3a'); D.h('3.1.1 Mapas de Empatía', 3)
    D.p('Los mapas de empatía organizan lo que cada usuario ve, escucha, piensa, siente, dice y hace, junto con sus dolores y ganancias. Complementan el lienzo del meta-usuario de la Tabla 5. El del productor se basa en el trabajo de campo con 20 productores; el del consumidor es una hipótesis del equipo que debe confirmarse con las entrevistas del Apéndice B.')
    D.table('Mapa de empatía del productor de café', ['Dimensión', 'Contenido'], [
        ['Ve', 'Su parcela con poca nutrición, plagas, café mezclado al cosechar, precios que cambian sin explicación.'],
        ['Escucha', 'Del acopiador, el precio del día; de otros productores, quién paga más; de los técnicos, prácticas que no puede costear.'],
        ['Piensa y siente', 'Preocupación por pagar las deudas de la campaña; desconfianza hacia quien promete mejores precios; orgullo por su café.'],
        ['Dice y hace', 'Vende al acopiador más cercano; seca en patio o plásticos; acepta visitas si traen ingreso (85%).'],
        ['Dolores', 'Falta de capital para abonar; roya y broca; mezcla de calidades; poca capacidad de negociar el precio.'],
        ['Ganancias', 'Producir más y mejor, vender a un comprador seguro con prima y recibir apoyo técnico sin endeudarse.'],
    ], note='Adaptado del trabajo de campo del equipo (Apéndice A).', widths=[1.3, 5.2])
    D.table('Mapa de empatía del consumidor de café de especialidad (hipótesis)', ['Dimensión', 'Contenido'], [
        ['Ve', 'Marcas con historias de origen en redes, cafeterías de especialidad, bolsas con datos de finca y variedad.'],
        ['Escucha', 'Recomendaciones de baristas, amigos y comunidades de café; críticas a café viejo o mezclado.'],
        ['Piensa y siente', 'Quiere un café fresco y auténtico; duda de que el origen sea real; le gusta ayudar si ve el resultado.'],
        ['Dice y hace', 'Compra café en grano una vez al mes o más; compara precios; comparte lo que compra y lo regala.'],
        ['Dolores', 'Desconfianza en el origen, café no siempre fresco, poca conexión con el productor (Tabla 5).'],
        ['Ganancias', 'Café fresco y trazable, impacto visible en el productor y una experiencia que puede compartir.'],
    ], note='Elaboración propia a partir del lienzo del meta-usuario (Tabla 5). Pendiente de validar con la encuesta del Apéndice B.', widths=[1.3, 5.2])
    D.at('c3b'); D.h('3.2.1 Mapa de Experiencia del Consumidor', 3)
    D.p('El mapa de la Tabla 8 describe solo al productor. La Tabla 48 presenta el recorrido del consumidor con Cosecha Viva, de modo que cada etapa tenga una oportunidad de mejora y una métrica que se mide en el piloto.')
    D.table('Mapa de experiencia del consumidor con Cosecha Viva', ['Etapa', 'Qué hace', 'Qué siente', 'Punto de dolor', 'Oportunidad', 'Métrica'], [
        ['Descubre', 'Ve un video o una recomendación y entra a la página.', 'Curiosidad', 'No sabe si el origen es real.', 'Video de la finca y ficha del productor.', 'Visitas calificadas'],
        ['Decide', 'Compara el paquete de S/ 140 con otros cafés.', 'Duda', 'Pagar por adelantado sin ver el producto.', 'Garantía de entrega, fecha de tueste y devolución.', 'Conversión (meta 3%)'],
        ['Compra', 'Elige productor, paga con tarjeta o billetera digital.', 'Expectativa', 'Pago complicado o sin comprobante claro.', 'Pago en pocos pasos y confirmación inmediata.', 'Abandono del pago'],
        ['Espera', 'Recibe videos por etapa del café.', 'Conexión', 'Que no haya novedades durante meses.', 'Un video por etapa y aviso de cada avance.', 'Apertura de mensajes'],
        ['Recibe', 'Abre la caja y escanea el código QR.', 'Emoción', 'Retrasos o café fuera de punto.', 'Entrega en fecha y tueste reciente.', 'Entregas a tiempo; satisfacción'],
        ['Disfruta y comparte', 'Prepara el café y comparte el video.', 'Orgullo', 'No sabe cómo preparar el grano.', 'Guía de preparación en el QR.', 'Recomendaciones'],
        ['Recompra y visita', 'Renueva su paquete o reserva el tour.', 'Lealtad', 'Falta de beneficios para quien repite.', 'Beneficio de recompra y reserva de visita.', 'Recompra 60%; tours 5%'],
    ], note='Elaboración propia. Los puntos de dolor son hipótesis; se confirman con las entrevistas a clientes del piloto (H5, H7 y H8).', widths=[0.8, 1.2, 0.7, 1.2, 1.4, 1.2], size=8.5)

    # ---------------- Capítulo IV ----------------
    D.pb(); D.h('Capítulo IV: Diseño del Producto o Servicio (sección nueva)', 1)
    D.where('Capítulo IV, dentro de 4.4 Propuesta de Valor, antes de la Tabla 12 o en su lugar como ampliación.')
    D.at('c4'); D.h('4.4.1 Lienzo de la Propuesta de Valor', 3)
    D.p('El lienzo de la propuesta de valor de Osterwalder et al. (2014) separa el perfil del cliente (trabajos, dolores y ganancias) del mapa de valor (productos y servicios, aliviadores de dolor y creadores de ganancia). Se construyó uno por cada lado del modelo. El encaje de la Tabla 12 resume la relación entre ambos.')
    D.table('Lienzo de la propuesta de valor: productor', ['Perfil del cliente', 'Mapa de valor'], [
        ['Trabajos: producir café para sostener a su familia; vender la cosecha; pagar la campaña; cuidar la parcela.', 'Productos y servicios: asistencia técnica, kit de fertirriego y secado en comodato, compra del café selecto, visitas a la finca.'],
        ['Dolores: falta de capital para abonar; roya y broca; mezcla de calidades; precio fijado por el intermediario.', 'Aliviadores: el kit se entrega sin costo inicial; el técnico enseña cosecha selectiva y manejo de plagas; el precio y el piso están pactados por escrito.'],
        ['Ganancias: más producción y mejor precio; un comprador seguro; reconocimiento de su comunidad.', 'Creadores de ganancia: prima de 33% sobre el intermediario (S/ 20 frente a S/ 15 por kg); mayor rendimiento; ingreso adicional por visitas.'],
    ], note='Elaboración propia a partir del trabajo de campo (Apéndice A) y de la Tabla 12. Adaptado de Osterwalder et al. (2014).', widths=[3.25, 3.25])
    D.table('Lienzo de la propuesta de valor: consumidor', ['Perfil del cliente', 'Mapa de valor'], [
        ['Trabajos: tomar un buen café en casa; regalar algo con historia; apoyar a productores; aprender sobre el origen.', 'Productos y servicios: paquete de cosecha de 1 kg (cuatro bolsas de 250 g) por preventa; videos por etapa; código QR por lote; tour a la finca.'],
        ['Dolores: desconfianza en el origen, café no siempre fresco, poca conexión con el productor.', 'Aliviadores: tueste cercano a la entrega, fecha de tueste en el QR, ficha del productor y evidencia en video.'],
        ['Ganancias: café fresco y trazable, impacto visible y una experiencia que puede compartir.', 'Creadores de ganancia: acompañar la cosecha, conocer al productor y visitar la finca; el pago por adelantado financia la tecnificación.'],
    ], note='Elaboración propia. Pendiente de validar con la encuesta del Apéndice B. Adaptado de Osterwalder et al. (2014).', widths=[3.25, 3.25])
    D.p('Referencia nueva para la lista: Osterwalder, A., Pigneur, Y., Bernarda, G., & Smith, A. (2014). Value proposition design: How to create products and services customers want. John Wiley & Sons.', italic=True, size=10)

    # ---------------- Capítulo V ----------------
    D.pb(); D.h('Capítulo V: Modelo de Negocio (secciones nuevas)', 1)
    D.where('Capítulo V: 5.3.1 va dentro de 5.3 Escalabilidad; 5.5 y 5.6 van después de 5.4 Sostenibilidad. Actualice el índice.')
    D.at('c5a'); D.h('5.3.1 Hoja de Ruta de Escalamiento y Estrategia de Salida', 3)
    D.p('El modelo crece por cohortes de productores y por clientes. La Tabla 51 muestra la escala que alcanza cada año según el modelo financiero y las condiciones que deben cumplirse antes de pasar a la siguiente etapa.')
    D.table('Hoja de ruta de escalamiento', ['Etapa', 'Año', 'Paquetes', 'Clientes activos', 'Productores', 'Condición para avanzar'], [
        ['Piloto', 'Meses 1 a 6', '300', '—', '10 fincas', 'H5 y H6 cumplidas; VAN recalculado positivo.'],
        ['Consolidación en Chanchamayo', '1', '12,000', '2,000', '33', 'Cobertura de abastecimiento de 100%; conversión cercana a 3%.'],
        ['Crecimiento', '2 a 3', '26,000 a 40,000', '4,333 a 6,667', '49 a 77', 'Retención cercana a 60% y LTV/CAC mayor que 3.'],
        ['Expansión de oferta', '4 a 5', '54,000 a 66,000', '9,000 a 11,000', '96 a 117', 'Replicar el modelo en Satipo y Villa Rica con la misma disciplina de calidad.'],
    ], note='Adaptado del modelo financiero (hojas Demanda y FCL). El piloto no forma parte de los cinco años proyectados.', widths=[1.3, 0.8, 1.0, 1.0, 0.9, 1.7])
    D.p('La estrategia de salida es una opción de los socios y no está incluida en el flujo de caja, que liquida los activos al quinto año sin perpetuidad. Las alternativas son: continuar la operación; vender la participación a una tostaduría o comercializadora de café; incorporar un inversionista con aporte de capital para escalar a otras regiones o cultivos; o transferir la plataforma de trazabilidad como servicio a otras marcas. La decisión dependerá de la retención de clientes, de la fidelidad de los productores y del VAN recalculado con los datos del piloto.')

    D.at('c5b'); D.h('5.5 Plan de Marketing', 2)
    D.p('El plan de marketing define a quién se vende, qué se promete, por qué canales y con qué presupuesto. Sus cifras provienen del modelo financiero (hoja Plan_Marketing del Excel): el primer año exige 2,000 clientes nuevos, equivalentes a 66,667 visitas calificadas con una conversión de 3% y un CAC de S/ 80.')
    D.h('5.5.1 Objetivos', 3)
    D.bullets(['Conseguir 300 paquetes de preventa con pago real en el piloto, con conversión mínima de 3% (H5).',
               'Alcanzar 12,000 paquetes y 2,000 clientes activos en el año 1, y 66,000 paquetes y 11,000 clientes en el año 5.',
               'Mantener un CAC menor o igual que S/ 80 y una relación LTV/CAC mayor que 3 (el modelo base da 9.91).',
               'Lograr 60% de recompra anual y que 5% de los clientes reserve un tour (H8).'])
    D.h('5.5.2 Segmentación, objetivo y posicionamiento', 3)
    D.p('El segmento objetivo es el consumidor de café de especialidad de Lima y la costa, de 25 a 55 años, que compra café al menos una vez al mes (sección 2.5). El posicionamiento es: "el café que usted financia antes de la cosecha y sigue hasta su taza, con un productor que gana más". La promesa tiene tres pruebas: videos por etapa, código QR por lote y visita opcional a la finca.')
    D.h('5.5.3 Mezcla de marketing', 3)
    D.table('Mezcla de marketing de Cosecha Viva (7P)', ['Elemento', 'Decisión'], [
        ['Producto', 'Paquete de cosecha de 1 kg (cuatro bolsas de 250 g) con código QR, videos por etapa y ficha del productor.'],
        ['Precio', 'S/ 140 con IGV (S/ 118.64 sin IGV), 17% más que la propuesta inicial de S/ 120 y cerca de la referencia de S/ 130 por kg del mercado digital. Pruebas con S/ 140 y S/ 130, con S/ 120 como precio de reserva. El VAN se anula con un precio de S/ 104.'],
        ['Plaza', 'Página propia de preventa con envío a domicilio en Lima y la costa; ventas directas para regalos corporativos.'],
        ['Promoción', 'Pauta digital, contenido de cosecha, referidos, alianzas con cafeterías y hoteles, y campañas por fechas.'],
        ['Personas', 'Ejecutivo de ventas y atención, comunicador audiovisual y técnicos que muestran la finca.'],
        ['Procesos', 'Preventa, pago, tueste por lotes con aliado, entrega en fecha y seguimiento posventa.'],
        ['Evidencia física', 'Bolsa con QR y fecha de tueste, videos, ficha del productor y tour a la finca.'],
    ], note='Elaboración propia.', widths=[1.3, 5.2])
    D.h('5.5.4 Canales, calendario y presupuesto', 3)
    D.p('La Tabla 54 propone la distribución de la inversión en adquisición por canal. Es una propuesta del equipo y se ajusta con el costo real por canal del piloto. La Tabla 55 relaciona las campañas con el ciclo del café.')
    D.table('Presupuesto de marketing según el modelo (S/)', ['Concepto', 'Año 1', 'Año 2', 'Año 3', 'Año 4', 'Año 5'], [
        ['Clientes nuevos a captar', '2,000', '3,133', '4,067', '5,000', '5,600'],
        ['Inversión en adquisición (CAC de S/ 80)', '160,000', '250,640', '325,360', '400,000', '448,000'],
        ['Marketing de retención y contenido (2%)', '28,475', '61,695', '94,915', '128,136', '156,610'],
        ['Total marketing', '188,475', '312,335', '420,275', '528,136', '604,610'],
        ['Marketing / ventas netas', '12.9%', '9.9%', '8.6%', '8.0%', '7.5%'],
    ], note='Adaptado del modelo financiero (hojas Demanda, EERR y Plan_Marketing).', widths=[2.6, 0.8, 0.8, 0.8, 0.8, 0.8])
    D.table('Distribución de la inversión en adquisición por canal (propuesta del equipo)', ['Canal', '% del total', 'Año 1 (S/)', 'Año 5 (S/)', 'Indicador'], [
        ['Pauta digital (Meta y Google)', '45%', '72,000', '201,600', 'Costo por visita y conversión'],
        ['Alianzas con tostadurías, cafeterías y hoteles', '15%', '24,000', '67,200', 'Clientes por alianza'],
        ['Contenido y videos de cosecha (orgánico)', '15%', '24,000', '67,200', 'Visitas orgánicas y suscriptores'],
        ['Regalos corporativos y ventas B2B', '15%', '24,000', '67,200', 'Paquetes por contrato'],
        ['Programa de referidos', '10%', '16,000', '44,800', 'Clientes por referido'],
    ], note='Elaboración propia a partir de la hoja Plan_Marketing del modelo.', widths=[2.4, 0.8, 0.9, 0.9, 1.5])
    D.table('Calendario de campañas', ['Periodo', 'Campaña', 'Acción principal'], [
        ['Meses 1 a 2', 'Preparación', 'Acuerdos con productores; fichas y videos de tres productores; página de preventa.'],
        ['Meses 3 a 4', 'Prueba de preventa (H5)', 'Pauta digital para unas 10,000 visitas calificadas; prueba de precio de S/ 140 y S/ 130.'],
        ['Meses 5 a 9', 'Campaña de preventa del primer año', 'Preventa con pago anticipado; videos de la cosecha selectiva; reservas de tours.'],
        ['Meses 7 a 12', 'Entregas y recompra', 'Envío del café tostado, medición de satisfacción y oferta de recompra.'],
        ['Fechas de regalo', 'Regalos', 'Ofertas para fechas de regalo, como el Día de la Madre y Navidad; confirmar el calendario con los datos del piloto.'],
    ], note='Adaptado del cronograma de la Tabla 34.', widths=[1.1, 2.0, 3.4])
    D.h('5.5.5 Indicadores del plan', 3)
    D.table('Indicadores de marketing y metas', ['Indicador', 'Meta', 'Fuente'], [
        ['Visitas calificadas', '66,667 en el año 1', 'Plataforma de preventa'],
        ['Conversión', '3% o más', 'Plataforma de preventa'],
        ['CAC', 'S/ 80 o menos', 'Pauta y ventas'],
        ['Costo por visita calificada', 'S/ 2.40 o menos', 'Pauta'],
        ['Retención anual', '60% o más', 'Base de clientes'],
        ['LTV / CAC', 'Mayor que 3', 'Modelo financiero'],
        ['Satisfacción del cliente', '4 de 5 o más', 'Encuesta posentrega'],
    ], note='Adaptado del modelo financiero y de las hipótesis H5 a H8.', widths=[2.5, 2.0, 2.0])
    D.p('Advertencia sobre el presupuesto del piloto. La Tabla 32 destina S/ 6,000 de pauta a unas 10,000 visitas calificadas, es decir, S/ 0.60 por visita. El modelo supone S/ 2.40 por visita (CAC de S/ 80 con conversión de 3%), cuatro veces más. Con el costo del modelo, esas visitas costarían S/ 24,000. Se recomienda presupuestar el piloto con el costo del modelo o medir el costo real por visita antes de proyectar el año 1.')

    D.h('5.6 Aspectos Legales, Societarios y Regulatorios', 2)
    D.p('La Tabla 57 resume las obligaciones principales. Está elaborada como guía del equipo y debe revisarse con un asesor legal y tributario antes de constituir la empresa.')
    D.table('Marco legal y regulatorio del modelo', ['Aspecto', 'Medida del modelo', 'Responsable'], [
        ['Constitución', 'Sociedad anónima cerrada con los cuatro socios; inscripción en Registros Públicos y obtención del RUC.', 'Gerente general y asesor legal'],
        ['Régimen tributario', 'Régimen general del impuesto a la renta (29.5%, como en el modelo); comprobantes electrónicos.', 'Finanzas y contador externo'],
        ['Acuerdo con el productor', 'Contrato escrito de compraventa de café y de comodato del kit, con volumen, calidad mínima, precio con piso, autorización de imagen y causales de terminación.', 'Gerente general y asesor legal'],
        ['Producción y tueste', 'El tueste se terceriza con una tostaduría que cuenta con sus autorizaciones sanitarias; se exige copia de ellas en el contrato.', 'Operaciones'],
        ['Venta a distancia y preventa', 'Información clara de precio, fecha de entrega y devolución; libro de reclamaciones virtual; cumplimiento de las normas de protección al consumidor.', 'Gestión comercial'],
        ['Datos personales', 'Política de privacidad, consentimiento para el uso de datos de clientes y de la imagen de los productores.', 'Gestión comercial y asesor legal'],
        ['Marca', 'Registro de la marca Cosecha Viva y de su logotipo ante el Indecopi.', 'Gerente general'],
    ], note='Elaboración propia. Normas específicas y plazos por confirmar con el asesor legal.', widths=[1.4, 3.7, 1.4])
