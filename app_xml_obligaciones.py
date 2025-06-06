# -*- coding: utf-8 -*-
"""
Created on Tue May 20 14:34:03 2025

@author: vbarahona
"""

# %%importar librerias
import streamlit as st
import xml.etree.ElementTree as ET
from datetime import date,datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
from decimal import Decimal
import tempfile
import openpyxl
#background-color: rgb(255, 255, 255);
#background-color: rgb(171 , 190 , 76)
#background-color: rgb(120, 154, 61);
# Fondo personalizado y fuente
st.markdown("""
<style>
    body {
        background-color:rgb(171 , 190 , 76);
        font-family: 'Handel Gothic', 'Frutiger light - Roman';
    }
    .stApp {
        background-color: rgb(255, 255, 255);
        font-family: 'Frutiger Bold', sans-serif;
    }
</style>
    """, unsafe_allow_html=True)
 
# Logo a la izquierda y título a la derecha
col1, col2 = st.columns([1, 2])
with col1:
    st.image('https://www.finagro.com.co/sites/default/files/logo-front-finagro.png', width=200)
with col2:
    st.markdown(
        '<h1 style="color: rgb(120,154,61); font-size: 2.25rem; font-weight: bold;">Generador de XML de Obligaciones a partir de un archivo Excel</h1>',
        unsafe_allow_html=True
    )

# Columnas predeterminadas para el archivo Excel
required_columns = [
'Tipo_de_cartera', 'Codigo_intermediario','Codigo_de_programa',
'Tipo_de_productor',
'Numero_del_pagare',
'Fecha_de_suscripcion',
'Ciudad_de_Inversion',
'Identificacion_del_primer_beneficiario',
'Tipo_de_Identificacion',
'Nombre_del_beneficiario_o_razon_social',
'Email_Beneficiario',
'Telefono_Beneficiario',
'Fecha_de_activos',
'Monto_Activos',
'Direccion_Beneficiario',
'Plazo',
'Tipo_plan_pagos',
'Capital_total',
'Porcentaje_Fag',
'Indicativo_Fag',
'Tipo_Comision',
'Puntos_IBR',
'Ubicacion_Predio',
'Codigo_oficina_de_origen',
'Producto_relacionado',
'Codigo_destino_1',
'Unidades_destino_1',
'Costo_de_Inversión_destino_1',
'Valor_a_Financiar_destino_1',
'Codigo_destino_2',
'Unidades_destino_2',
'Costo_de_Inversión_destino_2',
'Valor_a_Financiar_destino_2',
'Codigo_destino_3',
'Unidades_destino_3',
'Costo_de_Inversión_destino_3',
'Valor_a_Financiar_destino_3',
'Codigo_destino_4',
'Unidades_destino_4',
'Costo_de_Inversión_destino_4',
'Valor_a_Financiar_destino_4',
'Valor_Ingresos',
'Fecha_Corte_Ingresos'
]

st.markdown(
    '<span style="color: rgb(120, 154, 61); font-size: 22px;">Validador de Columnas Requeridas</span>',
    unsafe_allow_html=True
)
xls_file = st.file_uploader("", type=["xlsx", "xls"])

if xls_file:
    df = pd.read_excel(xls_file, engine='openpyxl')
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        st.error("❌ Faltan las siguientes columnas en el archivo:")
        for col in missing_columns:
            st.markdown(f"- **{col}**")
    else:
        st.success("✅ Todas las columnas requeridas están presentes.")
        df = df.dropna(subset=['Numero_del_pagare'])
        df['Fecha_Corte_Ingresos'] = pd.to_datetime(df['Fecha_Corte_Ingresos'], format='%Y/%m/%d')
        df['Fecha_de_activos'] = pd.to_datetime(df['Fecha_de_activos'], format='%Y/%m/%d')
        df['Fecha_de_suscripcion'] = pd.to_datetime(df['Fecha_de_suscripcion'], format='%Y/%m/%d')

        df['Fecha_Corte_Ingresos'] = df['Fecha_Corte_Ingresos'].dt.strftime('%Y-%m-%d')
        df['Fecha_de_activos'] = df['Fecha_de_activos'].dt.strftime('%Y-%m-%d')
        df['Fecha_de_suscripcion'] = df['Fecha_de_suscripcion'].dt.strftime('%Y-%m-%d')

        Valor_creditos = str(sum(df['Capital_total'].astype('float64')))
        Cantidad_creditos = str(len(df))

        # Formulario de parámetros
        with st.form("form_parametros"):
            fecha_Desembolso_str = st.date_input("Fecha de desembolso", value=date.today())
            Codigo_de_programa = st.text_input("Código del programa", value="501")
            cod_intermediario = st.text_input("Código del intermediario", value="203018")
            tipo_plan_checkbox = st.checkbox("¿Es un plan de pagos tipo bullet?", key="tipo_plan_checkbox")
            tipo_plan = 1 if tipo_plan_checkbox else 0
            submitted = st.form_submit_button("Confirmar parámetros")

        if submitted:
            st.markdown("""
            <style>
                body {
                    background-color:rgb(255 , 255 , 255);
                    font-family: 'Handel Gothic', 'Frutiger light - Roman';
                }
                .stApp {
                    background-color: rgb(171, 190, 76);
                    font-family: 'Frutiger Bold', sans-serif;
                }
            </style>
                """, unsafe_allow_html=True)
            st.subheader("Resumen de datos ingresados:")
            st.write(f"Fecha de desembolso: {fecha_Desembolso_str.strftime('%Y-%m-%d')}")
            #st.write(f"Código del programa: {Codigo_de_programa}")
            st.write(f"Código del intermediario: {cod_intermediario}")
            st.write(f"Tipo de plan: {'Bullet' if tipo_plan == 1 else 'Cuotas capital simétricas'}")
            st.write(f"Cantidad de créditos: {Cantidad_creditos}")
            st.write(f"Valor total créditos: {sum(df['Capital_total'].astype('float64')):.2f}")

            #st.header("Generar XML", divider=True)
            #if st.button("Generar XML"):
                ### rest of code
            try:
                # XML generation logic here
                ET.register_namespace('', "http://www.finagro.com.co/sit")
                obligaciones = ET.Element("{http://www.finagro.com.co/sit}obligaciones",
                                          cifraDeControl=Cantidad_creditos,
                                          cifraDeControlValor=Valor_creditos)
                fecha_Desembolso = fecha_Desembolso_str.strftime('%Y-%m-%d')
                #fecha_Desembolso = date(2025, 5, 9) # indicar fecha desembolso
                #Codigo_de_programa = '126' # indicar código del programa
                #cod_intermediario = '203018' # indicar código del intermediario
                #tipo_plan = 0 # solo va 1 o cero | # si tipo_plan = 1 entonces bullet sino cuotas capital simétricas
                st.write(f"Tipo plan: {tipo_plan}")
                st.dataframe(df)
                for index, row in df.iterrows():
                    # Crear vencimiento final
                    fechaFinal = pd.to_datetime(row['Fecha_de_suscripcion'],format ='%Y-%m-%d') + relativedelta(months=int(row['Plazo'])) 
                    fechaFinal = fechaFinal.strftime('%Y-%m-%d')
                    # Crear el elemento 'obligacion'
                    obligacion = ET.SubElement(obligaciones, "{http://www.finagro.com.co/sit}obligacion",
                                               tipoCartera= row['Tipo_de_cartera'],
                                               programaCredito = str(row['Codigo_de_programa']),
                                               tipoOperacion="1",
                                               tipoMoneda="1",
                                               tipoAgrupamiento="1",
                                               numeroPagare= row['Numero_del_pagare'],
                                               numeroObligacionIntermediario= str(datetime.today().strftime("%Y%d%H%M%S%f")),
                                               fechaSuscripcion=str(row['Fecha_de_suscripcion'] ),
                                               fechaDesembolso=str(fecha_Desembolso))
                
                    # Crear el elemento 'intermediario'
                    intermediario = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}intermediario",
                                                   oficinaPagare=str(row['Codigo_oficina_de_origen']),
                                                   oficinaObligacion=str(row['Codigo_oficina_de_origen']),
                                                   codigo=str(row['Codigo_intermediario']))
                
                    # Crear el elemento 'beneficiarios'
                    beneficiarios = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}beneficiarios",
                                                   cantidad="1")
                
                    # Crear el elemento 'beneficiario'
                    beneficiario = ET.SubElement(beneficiarios, "{http://www.finagro.com.co/sit}beneficiario",
                                                 correoElectronico=str(row['Email_Beneficiario']),
                                                 tipoAgrupacion="1",
                                                 tipoPersona="1",
                                                 tipoProductor=str(row['Tipo_de_productor']),
                                                 actividadEconomica=str(row['Producto_relacionado']),
                                                 cumpleCondicionesProductorAgrupacion="true")
                
                    # Crear el elemento 'identificacion' dentro de 'beneficiario'
                    identificacion_beneficiario = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}identificacion",
                                                                tipo=str(row['Tipo_de_Identificacion']),
                                                                numeroIdentificacion=str(row['Identificacion_del_primer_beneficiario']))

                    
                    # Crear el elemento 'nombre' dentro de 'beneficiario'
                    #calcular por espacios
                    nombre_beneficiario = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}nombre",
                                                       primerNombre=row['Nombre_del_beneficiario_o_razon_social'],
                                                       segundoNombre="",
                                                       primerApellido="",
                                                       segundoApellido="",
                                                       Razonsocial="")
                
                    # Crear el elemento 'nombre' dentro de 'beneficiario'
                    direccionCorrespondencia = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}direccionCorrespondencia",
                                                    direccion="R|"+str(row['Direccion_Beneficiario']),
                                                    municipio=str(row['Ciudad_de_Inversion']))
                
                    # Crear el elemento 'nombre' dentro de 'beneficiario'
                    numeroTelefono = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}numeroTelefono",
                                                   prefijo="6",
                                                   numero=str(row['Telefono_Beneficiario']))
                
                    # Crear el elemento 'valorActivos' dentro de 'beneficiario'
                    valor_activos = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}valorActivos",
                                                    valor=str(row['Monto_Activos']),
                                                    fechaCorte=str(row['Fecha_de_activos']),
                                                    tipoDato="COP")
                
                    # Crear el elemento 'valorIngresos' dentro de 'beneficiario'
                    valor_ingresos = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}valorIngresos",
                                                     valor=str(row['Valor_Ingresos']),
                                                     fechaCorte=str(row['Fecha_Corte_Ingresos']),
                                                     tipoDato="COP")
                
                    # Crear el elemento 'proyecto'
                    proyecto = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}proyecto",
                                            fechaInicialEjecucion=str(fecha_Desembolso),
                                            fechaFinalEjecucion=str(fechaFinal))
                    # Se podrían agregar 'incentivo' y 'proyectosFinanciados' dentro de 'proyecto' si fuera necesario
                
                    # Crear el elemento 'predios'
                    predios = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}predios")
                
                    # Crear un elemento 'predio' dentro de 'predios'
                    predio = ET.SubElement(predios, "{http://www.finagro.com.co/sit}predio",
                                           tipo="1",
                                           municipio=row['Ciudad_de_Inversion'],
                                           direccion="R|" +str(row['Ubicacion_Predio']))
                
                    #pendiente crear loop a partir de "Indicativo Fag"
                    if row['Indicativo_Fag'] == "S":
                        # Crear el elemento 'garantiaFAG'
                        garantiaFAG = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}garantiaFAG",
                                                         tipoComision =str(row['Tipo_Comision']),
                                                         porcentajeCobertura = str(row['Porcentaje_Fag'])
                                                         )
                   
                    
                    # Crear el elemento 'destinosCredito'
                    destinos_credito = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}destinosCredito")
                
                    # pendiente loop a partir de la cantida de destinos, solo hay hasta 4 destinos   
                    # Crear un elemento 'destinoCredito' dentro de 'destinosCredito'
                    destino_credito = ET.SubElement(destinos_credito, "{http://www.finagro.com.co/sit}destinoCredito",
                                                    codigo=str(row['Codigo_destino_1']),
                                                    unidadesAFinanciar=str(row['Unidades_destino_1']),
                                                    costoInversion=str(row['Costo_de_Inversión_destino_1']))
                
                    # Crear el elemento 'destinoCreditoValorAFinanciar' dentro de 'destinoCredito'
                    destino_credito_valor = ET.SubElement(destino_credito, "{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar")
                    valor_a_financiar = ET.SubElement(destino_credito_valor, "{http://www.finagro.com.co/sit}valorAFinanciar", {"xmlns": ""})
                    valor_a_financiar.text=str(row['Valor_a_Financiar_destino_1'])
                    
                    if not row['Codigo_destino_2']!= row['Codigo_destino_2']: 
                        # Crear un elemento 'destinoCredito' dentro de 'destinosCredito'
                        destino_credito = ET.SubElement(destinos_credito, "{http://www.finagro.com.co/sit}destinoCredito",
                                                        codigo=str(row['Codigo_destino_2']),
                                                        unidadesAFinanciar=str(row['Unidades_destino_2']),
                                                        costoInversion=str(row['Costo_de_Inversión_destino_2']))
                
                        # Crear el elemento 'destinoCreditoValorAFinanciar' dentro de 'destinoCredito'
                        destino_credito_valor = ET.SubElement(destino_credito, "{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar")
                        valor_a_financiar = ET.SubElement(destino_credito_valor, "{http://www.finagro.com.co/sit}valorAFinanciar", {"xmlns": ""})
                        valor_a_financiar.text=str(row['Valor_a_Financiar_destino_2'])
                    
                    if not row['Codigo_destino_3']!= row['Codigo_destino_3']: 
                        # Crear un elemento 'destinoCredito' dentro de 'destinosCredito'
                        destino_credito = ET.SubElement(destinos_credito, "{http://www.finagro.com.co/sit}destinoCredito",
                                                        codigo=str(row['Codigo_destino_3']),
                                                        unidadesAFinanciar=str(row['Unidades_destino_3']),
                                                        costoInversion=str(row['Costo_de_Inversión_destino_3']))
                
                        # Crear el elemento 'destinoCreditoValorAFinanciar' dentro de 'destinoCredito'
                        destino_credito_valor = ET.SubElement(destino_credito, "{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar")
                        valor_a_financiar = ET.SubElement(destino_credito_valor, "{http://www.finagro.com.co/sit}valorAFinanciar", {"xmlns": ""})
                        valor_a_financiar.text=str(row['Valor_a_Financiar_destino_3'])
                    
                    if not row['Codigo_destino_4']!= row['Codigo_destino_3']: 
                        # Crear un elemento 'destinoCredito' dentro de 'destinosCredito'
                        destino_credito = ET.SubElement(destinos_credito, "{http://www.finagro.com.co/sit}destinoCredito",
                                                        codigo=str(row['Codigo_destino_4']),
                                                        unidadesAFinanciar=str(row['Unidades_destino_4']),
                                                        costoInversion=str(row['Costo_de_Inversión_destino_4']))
                
                        # Crear el elemento 'destinoCreditoValorAFinanciar' dentro de 'destinoCredito'
                        destino_credito_valor = ET.SubElement(destino_credito, "{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar")
                        valor_a_financiar = ET.SubElement(destino_credito_valor, "{http://www.finagro.com.co/sit}valorAFinanciar", {"xmlns": ""})
                        valor_a_financiar.text=str(row['Valor_a_Financiar_destino_4'])
                    
                    # Crear el elemento 'financiacion'
                    financiacion = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}financiacion",
                                                   fechaVencimientoFinal=str(fechaFinal),
                                                   plazoCredito=str(row['Plazo']),
                                                   valorTotalCredito=str(row['Capital_total']),
                                                   porcentaje="100",
                                                   valorObligacion=str(row['Capital_total']))
                
                    # Datos para las cuotas
                    datos_cuotas = []
                    cantidad_cuotas= int(int(row['Plazo'])/int(row['Tipo_plan_pagos']))
                    cuota_capital = int(int(row['Capital_total'])/cantidad_cuotas)
                    
                    ult_cuota_capital = cuota_capital if int(row['Capital_total']) - (cantidad_cuotas*cuota_capital) == 0 else str(int(row['Capital_total']) - (Decimal(cantidad_cuotas-1)*Decimal(cuota_capital)))
                    fHasta = pd.to_datetime(row['Fecha_de_suscripcion'],format ='%Y-%m-%d')
                    for i in range(cantidad_cuotas-1):
                        meses = int(row['Tipo_plan_pagos'])
                        fHasta = fHasta + relativedelta(months=meses)
                        cuotas = {
                                        "registro": str(i+1),
                                        "fechaAplicacionHasta":str(date(int(fHasta.strftime('%Y')),int(fHasta.strftime('%m')),10)),
                                        "conceptoRegistroCuota": "I" if tipo_plan == 1 else "K",
                                        "periodicidadIntereses": "PE",
                                        "periodicidadCapital": "" if tipo_plan == 1 else "PE",
                                        "tasaBaseBeneficiario": "5",
                                        "margenTasaBeneficiario": str(row['Puntos_IBR']),
                                        "valorCuotaCapital": "0" if tipo_plan == 1 else str(cuota_capital),
                                        "porcentajeCapitalizacionIntereses": "0.0",
                                        "margenTasaRedescuento": "0"
                                        }
                            
                        datos_cuotas.append(cuotas)
                    cuotas = {
                                "registro": str(cantidad_cuotas),
                                "fechaAplicacionHasta": str(fechaFinal),
                                "conceptoRegistroCuota": "K",
                                "periodicidadIntereses": "PE",
                                "periodicidadCapital": "PE",
                                "tasaBaseBeneficiario": "5",
                                "margenTasaBeneficiario":  str(row['Puntos_IBR']),
                                "valorCuotaCapital": str(row['Capital_total']) if tipo_plan == 1 else str(ult_cuota_capital),
                                "porcentajeCapitalizacionIntereses": "0.0",
                                "margenTasaRedescuento": "0"
                            }
                         
                    datos_cuotas.append(cuotas)
                    # Crear el elemento 'planPagos'
                    plan_pagos = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}planPagos")
                
                    # Iterar sobre los datos de las cuotas y crear un elemento 'registroCuota' para cada uno
                    for dato_cuota in datos_cuotas:
                        registro_cuota = ET.SubElement(plan_pagos, "{http://www.finagro.com.co/sit}registroCuota",
                                                       registro=str(dato_cuota["registro"]),
                                                       fechaAplicacionHasta=str(dato_cuota["fechaAplicacionHasta"]),
                                                       conceptoRegistroCuota=dato_cuota["conceptoRegistroCuota"],
                                                       periodicidadIntereses=dato_cuota["periodicidadIntereses"],
                                                       periodicidadCapital=dato_cuota["periodicidadCapital"],
                                                       tasaBaseBeneficiario=dato_cuota["tasaBaseBeneficiario"],
                                                       margenTasaBeneficiario=dato_cuota["margenTasaBeneficiario"],
                                                       valorCuotaCapital=dato_cuota.get("valorCuotaCapital"),  # Usamos .get() por si es opcional
                                                       porcentajeCapitalizacionIntereses=dato_cuota.get("porcentajeCapitalizacionIntereses"),
                                                       margenTasaRedescuento=dato_cuota.get("margenTasaRedescuento"),
                                                      )
                    
                
                # Crear el árbol XML
                def sanitize_element(element):
                    if element.text is not None and not isinstance(element.text, str):
                        element.text = str(element.text)
                    for key, value in element.attrib.items():
                        if not isinstance(value, str):
                            element.attrib[key] = str(value)
                    for child in element:
                        sanitize_element(child)
                
                sanitize_element(obligaciones)

                tree = ET.ElementTree(obligaciones)
                ET.indent(tree, space="  ", level=0)
        
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp:
                            tree.write(tmp.name, encoding="UTF-8", xml_declaration=True)
                            st.success("✅ XML de obligaciones nuevas generado exitosamente.")
                            with open(tmp.name, "rb") as f:
                                st.download_button("📥 Descargar XML de Obligaciones nuevas", f, file_name="Obligaciones.xml", mime="application/xml")
            except Exception as e:
                st.error(f"Ocurrió un error al generar el XML: {e}")
