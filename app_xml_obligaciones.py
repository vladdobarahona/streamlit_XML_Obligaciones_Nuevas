# -*- coding: utf-8 -*-
"""
Created on Tue May 20 14:34:03 2025

@author: vbarahona
"""

# %%importar librerias
import streamlit as st
import xml.etree.ElementTree as ET
from datetime import date
from dateutil.relativedelta import relativedelta
import pandas as pd
from decimal import Decimal
import tempfile
import openpyxl

# Fondo personalizado y fuente
st.markdown("""
<style>
    body {
        background-color:rgb (171 , 190 , 76);
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
    'Tipo de Cartera', 'Tipo de Productor', 'Número de Pagare', 'Número de Pagare Anterior',
    'Fecha de Suscripción', 'Ciudad de Inversión', 'Numero de Identificacion',
    'Tipo Identificacion Finagro', 'Nombre Razón Social', 'Email Beneficiario',
    'Teléfono Beneficiario', 'Fecha de Activos', 'Monto Activos', 'Dirección Beneficiario',
    'Plazo', 'Tipo Plan de Pagos', 'Capital Total', 'Porcentaje Fag', 'Indicativo Fag',
    'Tipo Comisión', 'Puntos IBR', 'Ubicación Predio', 'Código Oficina', 'Producto Relacionado',
    'Código Destino 1', 'Unidades Destino 1', 'Costo Inversión 1', 'Valor a Financiar 1',
    'Código Destino 2', 'Unidades Destino 2', 'Costo Inversión 2', 'Valor a Financiar 2',
    'Código Destino 3', 'Unidades Destino 3', 'Costo Inversión 3', 'Valor a Financiar 3',
    'Código Destino 4', 'Unidades Destino 4', 'Costo Inversión 4', 'Valor a Financiar 4',
    'Valor Ingresos', 'Fecha de Ingresos', 'Moneda Ingresos', 'Moneda de Activos'
]

xls_file = st.file_uploader("", type=["xlsx", "xls"])

st.markdown(
    '<span style="color: rgb(120, 154, 61); font-size: 44px;">Validador de Columnas Requeridas</span>',
    unsafe_allow_html=True
)


if xls_file:
    df = pd.read_excel(xls_file, engine='openpyxl')
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        st.error("❌ Faltan las siguientes columnas en el archivo:")
        for col in missing_columns:
            st.markdown(f"- **{col}**")
    else:
        st.success("✅ Todas las columnas requeridas están presentes.")
        df = df.dropna(subset=['Número de Pagare'])
        df['Fecha de Ingresos'] = pd.to_datetime(df['Fecha de Ingresos'], format='%Y/%m/%d')
        df['Fecha de Activos'] = pd.to_datetime(df['Fecha de Activos'], format='%Y%m%d')
        df['Fecha de Suscripción'] = pd.to_datetime(df['Fecha de Suscripción'], format='%Y%m%d')

        df['Fecha de Ingresos'] = df['Fecha de Ingresos'].dt.strftime('%Y-%m-%d')
        df['Fecha de Activos'] = df['Fecha de Activos'].dt.strftime('%Y-%m-%d')
        df['Fecha de Suscripción'] = df['Fecha de Suscripción'].dt.strftime('%Y-%m-%d')

        Valor_creditos = str(sum(df['Capital Total'].astype('float64')))
        Cantidad_creditos = str(len(df))

        # Formulario de parámetros
        with st.form("form_parametros"):
            fecha_Desembolso_str = st.date_input("Fecha de desembolso", value=date.today())
            cod_programa = st.text_input("Código del programa", value="501")
            cod_intermediario = st.text_input("Código del intermediario", value="203018")
            tipo_plan_checkbox = st.checkbox("¿Es un plan de pagos tipo bullet?", key="tipo_plan_checkbox")
            tipo_plan = 1 if tipo_plan_checkbox else 0
            submitted = st.form_submit_button("Confirmar parámetros")

        if submitted:
            st.subheader("Resumen de datos ingresados:")
            st.write(f"Fecha de desembolso: {fecha_Desembolso}")
            st.write(f"Código del programa: {cod_programa}")
            st.write(f"Código del intermediario: {cod_intermediario}")
            st.write(f"Tipo de plan: {'Bullet' if tipo_plan == 1 else 'Cuotas capital simétricas'}")

            # Crear XML
            obligaciones = ET.Element("{http://www.finagro.com.co/sit}obligaciones",
                         cifraDeControl="1",
                         cifraDeControlValor="10000000.00")

            # Crear el elemento 'obligacion'
            obligacion = ET.SubElement(obligaciones, "{http://www.finagro.com.co/sit}obligacion",
                                       tipoCartera="1",
                                       programaCredito="126",
                                       tipoOperacion="1",
                                       tipoMoneda="1",
                                       tipoAgrupamiento="1",
                                       numeroPagare="00654526555245",
                                       numeroObligacionIntermediario="24512521542",
                                       fechaSuscripcion=str(date(2025, 5, 9)),
                                       fechaDesembolso=str(date(2025, 5, 9)))
            
            # Crear el elemento 'intermediario'
            intermediario = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}intermediario",
                                           oficinaPagare="1",
                                           oficinaObligacion="1",
                                           codigo="203018")
            
            # Crear el elemento 'beneficiarios'
            beneficiarios = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}beneficiarios",
                                           cantidad="1")
            
            # Crear el elemento 'beneficiario'
            beneficiario = ET.SubElement(beneficiarios, "{http://www.finagro.com.co/sit}beneficiario",
                                         correoElectronico="beneficiario@finagro.com.co",
                                         tipoAgrupacion="1",
                                         tipoPersona="1",
                                         tipoProductor="34",
                                         actividadEconomica="245100",
                                         cumpleCondicionesProductorAgrupacion="true")
            
            # Crear el elemento 'identificacion' dentro de 'beneficiario'
            identificacion_beneficiario = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}identificacion",
                                                        tipo="2",
                                                        numeroIdentificacion="1001635950")
            
            # Se podría agregar 'negocioFiduciario' dentro de 'identificacion_beneficiario' si fuera necesario
            
            # Crear el elemento 'nombre' dentro de 'beneficiario'
            nombre_beneficiario = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}nombre",
                                               primerNombre="MONSALVE",
                                               segundoNombre="TERRANO",
                                               primerApellido="JORGE",
                                               segundoApellido="ANDRES",
                                               Razonsocial="")
            
            # Crear el elemento 'nombre' dentro de 'beneficiario'
            direccionCorrespondencia = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}direccionCorrespondencia",
                                            direccion="R|VEREDA LA CONCHA, CARNICEROS, FINCA EL CORTIJO",
                                            municipio="63212")
            
            # Crear el elemento 'nombre' dentro de 'beneficiario'
            numeroTelefono = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}numeroTelefono",
                                           prefijo="6",
                                           numero="3187729283")
            
            # Crear el elemento 'valorActivos' dentro de 'beneficiario'
            valor_activos = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}valorActivos",
                                            valor="10000000.00",
                                            fechaCorte=str(date(2025, 2, 20)),
                                            tipoDato="Declarado")
            
            # Crear el elemento 'valorIngresos' dentro de 'beneficiario'
            valor_ingresos = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}valorIngresos",
                                             valor="5000.00",
                                             fechaCorte=str(date(2025, 2, 10)),
                                             tipoDato="COP")
            
            # Crear el elemento 'proyecto'
            proyecto = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}proyecto",
                                    fechaInicialEjecucion=str(date(2025, 5, 9)),
                                    fechaFinalEjecucion=str(date(2025, 6, 9)))
            # Se podrían agregar 'incentivo' y 'proyectosFinanciados' dentro de 'proyecto' si fuera necesario
            
            # Crear el elemento 'predios'
            predios = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}predios")
            
            # Crear un elemento 'predio' dentro de 'predios'
            predio = ET.SubElement(predios, "{http://www.finagro.com.co/sit}predio",
                                   tipo="1",
                                   municipio="63212",
                                   direccion="R|VEREDA LA CONCHA, CARNICEROS, FINCA EL CORTIJO")
            
            # Crear el elemento 'destinosCredito'
            destinos_credito = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}destinosCredito")
            
            # Crear un elemento 'destinoCredito' dentro de 'destinosCredito'
            destino_credito = ET.SubElement(destinos_credito, "{http://www.finagro.com.co/sit}destinoCredito",
                                            codigo="245100",
                                            unidadesAFinanciar="10.0",
                                            costoInversion="10000000")
            
            # Crear el elemento 'destinoCreditoValorAFinanciar' dentro de 'destinoCredito'
            destino_credito_valor = ET.SubElement(destino_credito, "{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar")
            valor_a_financiar = ET.SubElement(destino_credito_valor, "{http://www.finagro.com.co/sit}valorAFinanciar", {"xmlns": ""})
            valor_a_financiar.text="10000000"
            
            # Crear el elemento 'financiacion'
            financiacion = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}financiacion",
                                           fechaVencimientoFinal=str(date(2025, 6, 9)),
                                           plazoCredito="1",
                                           valorTotalCredito="10000000",
                                           porcentaje="100",
                                           valorObligacion="10000000")
            
            # Datos de ejemplo para las cuotas
            datos_cuotas = [
                {
                    "registro": 1,
                    "fechaAplicacionHasta": date(2025, 6, 20),
                    "conceptoRegistroCuota": "Pago Mensual",
                    "tasaBaseBeneficiario": "IBR",
                    "margenTasaBeneficiario": "2.5",
                    "valorCuotaCapital": "500.00",
                    "porcentajeCapitalizacionIntereses": "0.0",
                    "margenTasaRedescuento": "1.0",
                },
                {
                    "registro": 2,
                    "fechaAplicacionHasta": date(2025, 7, 20),
                    "conceptoRegistroCuota": "Pago Mensual",
                    "tasaBaseBeneficiario": "IBR",
                    "margenTasaBeneficiario": "2.5",
                    "valorCuotaCapital": "500.00",
                    "porcentajeCapitalizacionIntereses": "0.0",
                    "margenTasaRedescuento": "1.0",
                },
                # Puedes agregar más diccionarios para más cuotas
            ]
            
            # Crear el elemento 'planPagos'
            plan_pagos = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}planPagos")
            
            # Iterar sobre los datos de las cuotas y crear un elemento 'registroCuota' para cada uno
            for dato_cuota in datos_cuotas:
                registro_cuota = ET.SubElement(plan_pagos, "{http://www.finagro.com.co/sit}registroCuota",
                                               registro=str(dato_cuota["registro"]),
                                               fechaAplicacionHasta=str(dato_cuota["fechaAplicacionHasta"]),
                                               conceptoRegistroCuota=dato_cuota["conceptoRegistroCuota"],
                                               tasaBaseBeneficiario=dato_cuota["tasaBaseBeneficiario"],
                                               margenTasaBeneficiario=dato_cuota["margenTasaBeneficiario"],
                                               valorCuotaCapital=dato_cuota.get("valorCuotaCapital"),  # Usamos .get() por si es opcional
                                               porcentajeCapitalizacionIntereses=dato_cuota.get("porcentajeCapitalizacionIntereses"),
                                               margenTasaRedescuento=dato_cuota.get("margenTasaRedescuento"),
                                              )
            
            ############ Crear el elemento 'obligacion 2 ' #########################################################################################
            obligacion = ET.SubElement(obligaciones, "{http://www.finagro.com.co/sit}obligacion",
                                       tipoCartera="Comercial",
                                       programaCredito="Desarrollo Rural",
                                       tipoOperacion="Crédito",
                                       tipoMoneda="COP",
                                       tipoAgrupamiento="1",
                                       numeroPagare="PAG001",
                                       numeroObligacionIntermediario="OBI001",
                                       fechaSuscripcion=str(date(2025, 5, 15)),
                                       fechaDesembolso=str(date(2025, 5, 20)))
            
            # Crear el elemento 'intermediario'
            intermediario = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}intermediario",
                                           oficinaPagare="OF001",
                                           oficinaObligacion="OFB001",
                                           codigo="INT001")
            
            # Crear el elemento 'beneficiarios'
            beneficiarios = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}beneficiarios",
                                           cantidad="1")
            
            # Crear el elemento 'beneficiario'
            beneficiario = ET.SubElement(beneficiarios, "{http://www.finagro.com.co/sit}beneficiario",
                                         correoElectronico="beneficiario@example.com",
                                         tipoAgrupacion="Individual",
                                         tipoPersona="1",
                                         tipoProductor="Pequeño",
                                         cumpleCondicionesProductorAgrupacion="true",
                                         actividadEconomica="Agricultura")
            
            # Crear el elemento 'identificacion' dentro de 'beneficiario'
            identificacion_beneficiario = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}identificacion",
                                                        tipo="CC",
                                                        numeroIdentificacion="1234567890")
            # Se podría agregar 'negocioFiduciario' dentro de 'identificacion_beneficiario' si fuera necesario
            
            # Crear el elemento 'nombre' dentro de 'beneficiario'
            nombre_beneficiario = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}nombre",
                                               primerNombre="Juan",
                                               primerApellido="Pérez")
            
            # Crear el elemento 'valorActivos' dentro de 'beneficiario'
            valor_activos = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}valorActivos",
                                            valor="10000.00",
                                            fechaCorte=str(date(2025, 5, 10)),
                                            tipoDato="Declarado")
            
            # Crear el elemento 'valorIngresos' dentro de 'beneficiario'
            valor_ingresos = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}valorIngresos",
                                             valor="5000.00",
                                             fechaCorte=str(date(2025, 5, 10)),
                                             tipoDato="Declarado")
            
            # Crear el elemento 'proyecto'
            proyecto = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}proyecto")
            # Se podrían agregar 'incentivo' y 'proyectosFinanciados' dentro de 'proyecto' si fuera necesario
            
            # Crear el elemento 'predios'
            predios = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}predios")
            
            # Crear un elemento 'predio' dentro de 'predios'
            predio = ET.SubElement(predios, "{http://www.finagro.com.co/sit}predio",
                                   tipo="1",
                                   municipio="Bogotá",
                                   direccion="Calle 1 # 2-3")
            
            # Crear el elemento 'destinosCredito'
            destinos_credito = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}destinosCredito")
            
            # Crear un elemento 'destinoCredito' dentro de 'destinosCredito'
            destino_credito = ET.SubElement(destinos_credito, "{http://www.finagro.com.co/sit}destinoCredito",
                                            codigo="DC001",
                                            unidadesAFinanciar="10.0",
                                            costoInversion="2000.00")
            
            # Crear el elemento 'destinoCreditoValorAFinanciar' dentro de 'destinoCredito'
            destino_credito_valor = ET.SubElement(destino_credito, "{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar")
            valor_a_financiar = ET.SubElement(destino_credito_valor, "{http://www.finagro.com.co/sit}valorAFinanciar")
            valor_a_financiar.text = "1500.00"
            
            # Crear el elemento 'financiacion'
            financiacion = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}financiacion",
                                           fechaVencimientoFinal=str(date(2026, 5, 20)),
                                           plazoCredito="12",
                                           valorTotalCredito="10000.00",
                                           valorObligacion="10000.00")
            
            # Datos de ejemplo para las cuotas
            datos_cuotas = [
                {
                    "registro": 1,
                    "fechaAplicacionHasta": date(2025, 6, 20),
                    "conceptoRegistroCuota": "Pago Mensual",
                    "tasaBaseBeneficiario": "IBR",
                    "margenTasaBeneficiario": "2.5",
                    "valorCuotaCapital": "500.00",
                    "porcentajeCapitalizacionIntereses": "0.0",
                    "margenTasaRedescuento": "1.0",
                },
                {
                    "registro": 2,
                    "fechaAplicacionHasta": date(2025, 7, 20),
                    "conceptoRegistroCuota": "Pago Mensual",
                    "tasaBaseBeneficiario": "IBR",
                    "margenTasaBeneficiario": "2.5",
                    "valorCuotaCapital": "500.00",
                    "porcentajeCapitalizacionIntereses": "0.0",
                    "margenTasaRedescuento": "1.0",
                },
                # Puedes agregar más diccionarios para más cuotas
            ]
            
            # Crear el elemento 'planPagos'
            plan_pagos = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}planPagos")
            
            # Iterar sobre los datos de las cuotas y crear un elemento 'registroCuota' para cada uno
            for dato_cuota in datos_cuotas:
                registro_cuota = ET.SubElement(plan_pagos, "{http://www.finagro.com.co/sit}registroCuota",
                                               registro=str(dato_cuota["registro"]),
                                               fechaAplicacionHasta=str(dato_cuota["fechaAplicacionHasta"]),
                                               conceptoRegistroCuota=dato_cuota["conceptoRegistroCuota"],
                                               tasaBaseBeneficiario=dato_cuota["tasaBaseBeneficiario"],
                                               margenTasaBeneficiario=dato_cuota["margenTasaBeneficiario"],
                                               valorCuotaCapital=dato_cuota.get("valorCuotaCapital"),  # Usamos .get() por si es opcional
                                               porcentajeCapitalizacionIntereses=dato_cuota.get("porcentajeCapitalizacionIntereses"),
                                               margenTasaRedescuento=dato_cuota.get("margenTasaRedescuento"),
                                              )
            
            # Crear el árbol XML
            tree = ET.ElementTree(obligaciones)
                ET.indent(tree, space="  ", level=0)
        
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp:
                            tree.write(tmp.name, encoding="UTF-8", xml_declaration=True)
                            st.success("✅ XML de obligaciones nuevas generado exitosamente.")
                            with open(tmp.name, "rb") as f:
                                st.download_button("📥 Descargar XML de Obligaciones nuevas", f, file_name="Obligaciones.xml", mime="application/xml")
                                
            except Exception as e:
                st.error(f"Ocurrió un error al generar el XML: {e}")
