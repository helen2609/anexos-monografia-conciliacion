# Anexos de la monografia

```
================================================================================
PARA EL REVISOR - LEA PRIMERO
================================================================================
Para reproducir los resultados de la monografia NO necesita subir ninguna planilla
ni ningun PDF. Siga estos 3 pasos:

  1. Abra https://colab.research.google.com  ->  Archivo  ->  Subir cuaderno
     y elija 2_Codigo/reproducir_resultados_revisor.ipynb
  2. Menu  Entorno de ejecucion  ->  Ejecutar todas
  3. Espere unos 5 minutos. El cuaderno descarga solo los archivos del repositorio
     de GitHub y muestra las Tablas 4-3 a 4-6, la Tabla A-5 y las Figuras 4-4 a 4-8.
     Si no hubiera conexion con GitHub, el cuaderno pide subir tres archivos de este
     CD: 1_Datos/Datos_PowerBI.xlsx, 2_Codigo/modelos_v12.py y
     2_Codigo/analisis_estadistico.py.

Las cifras que debe obtener: LOF con AUC 0,676 (IC 95 % 0,53 a 0,81); reglas con
5 de 14 observadas y Fisher p = 0,358; LOF con 7 de 14 observadas en 119 ordenes.

El otro cuaderno, 2_Codigo/conciliacion_colab.ipynb, es la herramienta de uso
mensual de la empresa: pide la planilla real y los PDF del sistema del operador,
que no se incluyen por confidencialidad. No es necesario ejecutarlo para revisar
la monografia.
================================================================================

ANEXO PRINCIPAL - CONTENIDO DEL CD (version final)
Monografia: Deteccion de anomalias en la conciliacion de materiales para instalaciones de fibra optica
mediante reglas y aprendizaje automatico. Autora: Helen Angulo Laime.
La empresa contratista, el operador, los tecnicos (T01 a T27) y las solicitudes (O001 a O471) estan anonimizados.

CARPETAS
1_Datos/              Datos_PowerBI.xlsx: base anonimizada del analisis y del tablero. Hojas: Ordenes (471, con puntaje y
                      prioridad LOF), Diferencias (32), Tecnicos (27), Materiales (19), Modelos (14 enfoques), KPI (15).
                      Informe_conciliacion_anonimo.xlsx: informe de conciliacion generado por la herramienta.
                      Tablas_resultados/: CSV de todas las tablas de resultados y resumen_resultados.json.
2_Codigo/             reproducir_resultados_revisor.ipynb: reproduce el Capitulo 4 solo con la base anonimizada.
                      conciliacion_colab.ipynb: cuaderno de Google Colab (reglas, comparacion con PDF, exportacion y modelos).
                      modelos_v12.py: catorce enfoques, seleccion del numero de grupos, permutaciones, bootstrap y sensibilidad.
                      deteccion_anomalias.py: reglas de validacion e indice de riesgo.
                      comparador_pdf.py: comparacion de la planilla con los PDF del sistema del operador.
                      analisis_estadistico.py: Wilson, Fisher y Mann-Whitney para las reglas.
                      requirements.txt: librerias y versiones usadas.
3_Tablero_PowerBI/    Tablero_Conciliacion.pbix y medidas_DAX.txt.
4_Figuras/            Figuras del documento en PNG (numeradas como en la monografia).
5_Documento/          Monografia final en Word y PDF, y Resultados_reproducidos_revisor.pdf con la
                      ejecucion completa del cuaderno de reproduccion y sus resultados.

COMO REPRODUCIR LOS RESULTADOS DEL CAPITULO 4 (sin datos reales) - PARA EL REVISOR
Opcion A (recomendada): abrir 2_Codigo/reproducir_resultados_revisor.ipynb en Google Colab y elegir
  Entorno de ejecucion > Ejecutar todas. Descarga los archivos del repositorio de GitHub y muestra las tablas y
  figuras del Capitulo 4. Modo rapido: unos 5 minutos.
Opcion B (en una computadora con Python):
1. pip install -r 2_Codigo/requirements.txt
2. python 2_Codigo/modelos_v12.py 1_Datos/Datos_PowerBI.xlsx salida/
   Genera en salida/ las tablas (CSV) y figuras de los apartados 4.4 y del Anexo 5 (tarda unos 20 minutos por las
   pruebas de permutacion). Los resultados coinciden exactamente con el documento, porque los modelos no usan el
   tecnico ni el numero de solicitud.
   Nota: el script conserva su numeracion interna de archivos. Equivalencias con el documento:
   Tabla_4-3_comparacion_enfoques = Tabla 4-4; Tabla_4-4_carga_de_revision = Tabla 4-5;
   Tabla_4-5_perfil_segmentos = Tabla 4-3; Figura_4-4_comparacion_enfoques = Figura 4-6;
   Figura_4-5_curvas_ROC_PR = Figura 4-7; Figura_4-6_seleccion_grupos = Figura 4-4;
   Figura_4-7_perfil_grupos = Figura 4-5. La carpeta 4_Figuras ya usa la numeracion del documento.
3. python 2_Codigo/analisis_estadistico.py 1_Datos/Datos_PowerBI.xlsx  (pruebas de las reglas, apartado 4.6.3)

COMO USAR LA HERRAMIENTA CADA MES (con datos reales; uso interno de la empresa)
1. Abrir 2_Codigo/conciliacion_colab.ipynb en Google Colab y ejecutar las celdas en orden.
2. Subir la planilla de conciliacion (.xlsx), el ZIP con los PDF del sistema y, si existen, las observaciones del
   operador (observaciones_operador.csv).
3. El cuaderno genera el informe de conciliacion, la base del tablero y la prioridad LOF de cada orden.

NOTA: la planilla original, los PDF del sistema y las observaciones con numeros de solicitud reales no se incluyen por
contener informacion de la empresa y del operador.
```
