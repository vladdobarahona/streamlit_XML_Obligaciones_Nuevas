# 📄🔄 streamlit_XML_Obligaciones_Nuevas

Convertidor de archivo Excel a XML Obligaciones

Esta herramienta permite convertir un archivo Excel (con una estructura previamente definida) en un archivo XML, útil para la inscripción de operaciones en el aplicativo.

Para el uso adecuado, se deben informar cuatro (4) variables que se toman como parámetro para la creación del plan de pagos.

### 🧮 Tipos de plan de pagos

- **Bullet**: El capital del crédito se ubica en la última cuota calculada.
- **Lineal**: El capital total del crédito se distribuye simétricamente en la cantidad de cuotas establecidas.

### 📊 Estructura esperada del archivo Excel

La columna **[Tipo Plan de Pagos]** representa los meses entre cuotas:

- `1`: Mensual  
- `3`: Trimestral  
- `12`: Anual  
- etc.


🚀 Cómo ejecutar la aplicación

Carga el archivo Excel establecido

🧾 Campos del formulario

Fecha de desembolso: Selecciona la fecha en la que se realiza el desembolso.
Código del programa: Código identificador del programa financiero.
Código del intermediario: Código del intermediario financiero.

Tipo de plan:
✅ Marcado: Plan tipo bullet.
⬜ No marcado: Plan de cuotas de capital simétricas.

📦 Requisitos
Python 3.7 o superior
Streamlit
