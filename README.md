ANEXO PRINCIPAL - CONTENIDO DEL CD
Monografia: Deteccion de anomalias en la conciliacion de materiales de instalaciones FTTH y FTTB
mediante reglas y aprendizaje automatico. Autora: Helen Angulo Laime.
Los nombres de la empresa contratista y del operador, y de los tecnicos (T01 a T27), estan anonimizados.

CARPETAS
1_Datos/              Datos_PowerBI.xlsx : datos anonimizados usados en el analisis y en el tablero.
                      Hojas: Ordenes (471), Diferencias (32), Tecnicos (27), Materiales (19), Modelos (4), KPI (15).
                      Informe_conciliacion_anonimo.xlsx : informe de conciliacion generado por la herramienta
                      (Resumen, Ordenes_con_alerta, Diferencias_vs_sistema).
2_Codigo/             deteccion_anomalias.py : reglas de validacion, indice de riesgo y modelos.
                      comparador_pdf.py      : comparacion de la planilla con los reportes del sistema del operador.
                      conciliacion_colab.ipynb : cuaderno de Google Colab que ejecuta todo el proceso. Su ultima celda
                      ("Exportar datos para Power BI y figuras") genera Datos_PowerBI.xlsx y las Figuras 4-1 a 4-5.
3_Tablero_PowerBI/    Tablero_Conciliacion.pbix : tablero de dos paginas (Resumen y Detalle de ordenes).
                      medidas_DAX.txt : medidas DAX del tablero.
4_Figuras/            Figuras del documento (PNG, 300 dpi).
5_Documento/          Monografia en Word y PDF.

COMO REPRODUCIR LOS RESULTADOS
1. Abrir 2_Codigo/conciliacion_colab.ipynb en Google Colab (colab.research.google.com).
2. Ejecutar las celdas en orden. Cuando lo pida, subir la planilla de conciliacion mensual (.xlsx) y un archivo .zip con
   los PDF del sistema del operador.
3. El cuaderno genera el informe de conciliacion (Excel y texto) con el nivel de riesgo por orden y las diferencias.
   La ultima celda crea la carpeta salida_anexos/ con Datos_PowerBI.xlsx (tecnicos anonimizados) y las Figuras 4-1 a 4-5.
4. Para ver el tablero, abrir 3_Tablero_PowerBI/Tablero_Conciliacion.pbix en Power BI Desktop (Windows).
   Si pide origen de datos, apuntar a 1_Datos/Datos_PowerBI.xlsx.

NOTA: la planilla original y los PDF del sistema del operador no se incluyen por contener informacion de la empresa y del
operador. Los resultados con los datos de septiembre de 2026 son: 471 ordenes, 27 de riesgo ALTO, 92 de riesgo MEDIO,
32 diferencias contra el sistema en 23 ordenes (Bs 651,30 no facturados y Bs 366,36 facturados de mas).
