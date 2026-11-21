# Cambios en las cifras al subir el precio del paquete de S/ 120 a S/ 140

Solo cambió una celda del Excel: `Supuestos!B6` (de 120 a 140). Todo lo demás se recalcula con fórmulas.
Estas son las cifras de la tesis original (v4.1) que deben actualizarse.

| Cifra | Con S/ 120 | Con S/ 140 |
|---|---|---|
| Valor de venta sin IGV por paquete | S/ 101.69 | S/ 118.64 |
| Margen de contribución por paquete | S/ 39.13 (38.5%) | S/ 55.24 (46.6%) |
| Flujo de caja libre, año 0 | −S/ 285,509 | −S/ 271,700 |
| **Flujo de caja libre, año 1** | **−S/ 131,797** | **+S/ 48,543** |
| Flujo de caja libre, año 2 | S/ 218,343 | S/ 460,485 |
| Flujo de caja libre, año 3 | S/ 502,637 | S/ 947,356 |
| Flujo de caja libre, año 4 | S/ 794,391 | S/ 1,394,761 |
| Flujo de caja libre, año 5 | S/ 1,324,340 | S/ 2,044,317 |
| VAN financiero (COK 20.94%) | S/ 922,276 | S/ 2,061,174 |
| TIR | 70.4% | 128.8% |
| Relación beneficio-costo | 1.09 | 1.20 |
| Periodo de recupero descontado | 2.86 años | 1.74 años |
| VAN social al 8% (beneficios - costos, sin FCL ni inversión) | S/ 5,004,740 (método anterior, incorrecto) | **S/ 2,558,889** |
| LTV por cliente | S/ 556 | S/ 793 |
| LTV / CAC | 6.96 | 9.91 |
| Precio de quiebre del VAN | S/ 104.40 | S/ 104 (exacto 104.22) |
| Escenario pesimista | −S/ 303,224 | S/ 616,377 |
| Escenario estrés | −S/ 776,031 | S/ 98,728 |
| VAN esperado por escenarios | S/ 609,814 | S/ 1,713,197 |
| Probabilidad de VAN negativo (escenarios) | 30% | 0% |
| Marketing total, año 1 | S/ 184,407 | S/ 188,475 |
| Capital de trabajo | S/ 13,809 (año 0) | 0 (el anticipo de clientes cubre el inventario) |

Tablas de la tesis v4.1 afectadas: 16 (LTV), 19 a 28, 30 y 31 (VAN social), el Resumen Ejecutivo, el Abstract,
las conclusiones y el Apéndice F (supuesto de precio, fila 1).

**Advertencia.** S/ 140 es 8% más que la referencia de S/ 130 por kg del mercado digital y 40% más que los
S/ 100 de la encuesta de campo. El modelo no reduce la conversión por el mayor precio. La hipótesis H5 debe
probar S/ 140, S/ 130 y S/ 120.

## Corrección del VAN social
VAN social = beneficios sociales - costos sociales, descontados al 8%. No usa el flujo de caja libre ni la inversión.
Hojas nuevas del Excel: Beneficios Sociales, Costos Sociales y VAN Social. Beneficios: S/ 3,741,400 en valor presente; costos: S/ 1,182,511; relación beneficio-costo social: 3.16.
