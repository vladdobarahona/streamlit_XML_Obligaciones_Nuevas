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

st.markdown(
    '<span style="color: rgb(120, 154, 61); font-size: 22px;">Sube el archivo a convertir en XML (Excel)</span>',
    unsafe_allow_html=True
)

xls_file = st.file_uploader("", type=["xlsx"])

if xls_file:
    if st.button("Validar estructura"):
    # Subida de archivos
    
    
    #fecha_Desembolso = date(2025, 5, 9) # indicar fecha desembolso
    #cod_programa = '126' # indicar código del programa
    #cod_intermediario = '203018' # indicar código del intermediario
    #tipo_plan = 0 # solo va 1 o cero | # si tipo_plan = 1 entonces bullet sino cuotas capital simétricas
    
    #st.title("Parámetros de Desembolso")
    st.markdown(
            '<h1 style="color: rgb(120,154,61); font-size: 2.25rem; font-weight: bold;">"Parámetros de Desembolso</h1>',
            unsafe_allow_html=True
        )
    
    
    st.markdown(
        '<label style="color: rgb(11, 94, 94); font-weight: bold;">Indicar fecha de desembolso:</label>',
        unsafe_allow_html=True
    )
    
    # Fecha de desembolso
    fecha_Desembolso = st.date_input(
        label=" ",  # Empty label to avoid duplication
        value=date.today()
    )
    
    
    # Código del programa
    st.markdown(
        '<label style="color: rgb(11, 94, 94); font-weight: bold;">Indicar código del programa:</label>',
        unsafe_allow_html=True
    )
    
    cod_programa = st.text_input(
        label=" ",  # Etiqueta vacia para que no se duplique con el nombre antes indicado
        value="501"
    )
    
    # Código del intermediario
    st.markdown(
        '<label style="color: rgb(11, 94, 94); font-weight: bold;">Indicar código del intermediario:</label>',
        unsafe_allow_html=True
    )
    cod_intermediario = st.text_input(
        label=" ",  # Etiqueta vacia para que no se duplique con el nombre antes indicado
        value="203018"
    )
    
    # Tipo de plan
    st.markdown(
        '<label style="color: rgb(11, 94, 94); font-weight: bold;">¿Es un plan de pagos tipo bullet?</label>',
        unsafe_allow_html=True
    )
    tipo_plan_checkbox = st.checkbox(" ", key="tipo_plan_checkbox") # Empty label to avoid duplication
    tipo_plan = 1 if tipo_plan_checkbox else 0
    
    # Mostrar los valores ingresados
    st.subheader("Resumen de datos ingresados:")
    st.write(f"Fecha de desembolso: {fecha_Desembolso}")
    st.write(f"Código del programa: {cod_programa}")
    st.write(f"Código del intermediario: {cod_intermediario}")
    st.write(f"Tipo de plan: {'Bullet' if tipo_plan == 1 else 'Cuotas capital simétricas'}")
    
    
    xls_file = xls_file.dropna(subset='Número de Pagare')
    xls_file['Fecha de Ingresos']= pd.to_datetime(xls_file['Fecha de Ingresos'], format='%Y/%m/%d')
    xls_file['Fecha de Activos']= pd.to_datetime(xls_file['Fecha de Activos'], format='%Y%m%d')
    xls_file['Fecha de Suscripción'] = pd.to_datetime(xls_file['Fecha de Suscripción'],format='%Y%m%d')
    
    xls_file['Fecha de Ingresos']= xls_file['Fecha de Ingresos'].dt.strftime('%Y-%m-%d')
    xls_file['Fecha de Activos']= xls_file ['Fecha de Activos'].dt.strftime('%Y-%m-%d')
    xls_file['Fecha de Suscripción'] = xls_file['Fecha de Suscripción'].dt.strftime('%Y-%m-%d')
    
    valores_nulos = xls_file.isna().sum()
    
    Valor_creditos = str(sum(xls_file['Capital Total'].astype('float64')))
    Cantidad_creditos = str(len(xls_file))
    
    print("cantidad de registros varios:", valores_nulos)
    print(f"Usted ha cargado un archivo con {Cantidad_creditos} créditos por valor de {Valor_creditos:.2f}")

#%% Crear el elemento raíz 'obligaciones' con sus atributos
ET.register_namespace('', "http://www.finagro.com.co/sit")
obligaciones = ET.Element("{http://www.finagro.com.co/sit}obligaciones",
                         cifraDeControl=Cantidad_creditos,
                         cifraDeControlValor=Valor_creditos )

if xls_file:
    if st.button("Generar XML"):
        for index,row in xls_file.iterrows():
            # Crear vencimiento final
            fechaFinal = pd.to_datetime(row['Fecha de Suscripción'],format ='%Y-%m-%d') + relativedelta(months=int(row['Plazo'])) 
            fechaFinal = fechaFinal.strftime('%Y-%m-%d')
            # Crear el elemento 'obligacion'
            obligacion = ET.SubElement(obligaciones, "{http://www.finagro.com.co/sit}obligacion",
                                       tipoCartera= row['Tipo de Cartera'],
                                       programaCredito = cod_programa,
                                       tipoOperacion="1",
                                       tipoMoneda="1",
                                       tipoAgrupamiento="1",
                                       numeroPagare= row['Número de Pagare'],
                                       numeroObligacionIntermediario= row['Número de Pagare'],
                                       fechaSuscripcion=str(row['Fecha de Suscripción'] ),
                                       fechaDesembolso=str(fecha_Desembolso))
        
            # Crear el elemento 'intermediario'
            intermediario = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}intermediario",
                                           oficinaPagare=str(row['Código Oficina']),
                                           oficinaObligacion=str(row['Código Oficina']),
                                           codigo=cod_intermediario)
        
            # Crear el elemento 'beneficiarios'
            beneficiarios = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}beneficiarios",
                                           cantidad="1")
        
            # Crear el elemento 'beneficiario'
            beneficiario = ET.SubElement(beneficiarios, "{http://www.finagro.com.co/sit}beneficiario",
                                         correoElectronico=str(row['Email Beneficiario']),
                                         tipoAgrupacion="1",
                                         tipoPersona="1",
                                         tipoProductor=str(row['Tipo de Productor']),
                                         actividadEconomica=str(row['Producto Relacionado']),
                                         cumpleCondicionesProductorAgrupacion="true")
        
            # Crear el elemento 'identificacion' dentro de 'beneficiario'
            identificacion_beneficiario = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}identificacion",
                                                        tipo="2",
                                                        numeroIdentificacion=str(row['Numero de Identificacion']))
        
            # Se podría agregar 'negocioFiduciario' dentro de 'identificacion_beneficiario' si fuera necesario
        
            # Crear el elemento 'nombre' dentro de 'beneficiario'
            #calcular por espacios
            nombre_beneficiario = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}nombre",
                                               primerNombre=row['Nombre Razón Social'],
                                               segundoNombre="",
                                               primerApellido="",
                                               segundoApellido="",
                                               Razonsocial="")
        
            # Crear el elemento 'nombre' dentro de 'beneficiario'
            direccionCorrespondencia = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}direccionCorrespondencia",
                                            direccion="R|"+str(row['Ubicación Predio']),
                                            municipio=str(row['Ciudad de Inversión']))
        
            # Crear el elemento 'nombre' dentro de 'beneficiario'
            numeroTelefono = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}numeroTelefono",
                                           prefijo="6",
                                           numero=str(row['Teléfono Beneficiario']))
        
            # Crear el elemento 'valorActivos' dentro de 'beneficiario'
            valor_activos = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}valorActivos",
                                            valor=str(row['Monto Activos']),
                                            fechaCorte=str(row['Fecha de Activos']),
                                            tipoDato=str(row['Moneda de Activos']))
        
            # Crear el elemento 'valorIngresos' dentro de 'beneficiario'
            valor_ingresos = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}valorIngresos",
                                             valor=str(row['Valor Ingresos']),
                                             fechaCorte=str(row['Fecha de Ingresos']),
                                             tipoDato=str(row['Moneda Ingresos']))
        
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
                                   municipio=row['Ciudad de Inversión'],
                                   direccion="R|" +str(row['Ubicación Predio']))
        
            #pendiente crear loop a partir de "Indicativo Fag"
            if row['Indicativo Fag'] == "S":
                # Crear el elemento 'garantiaFAG'
                garantiaFAG = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}garantiaFAG",
                                                 tipoComision =str(row['Tipo Comisión']),
                                                 porcentajeCobertura = str(row['Porcentaje Fag'])
                                                 )
           
            
            # Crear el elemento 'destinosCredito'
            destinos_credito = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}destinosCredito")
        
            # pendiente loop a partir de la cantida de destinos, solo hay hasta 4 destinos   
            # Crear un elemento 'destinoCredito' dentro de 'destinosCredito'
            destino_credito = ET.SubElement(destinos_credito, "{http://www.finagro.com.co/sit}destinoCredito",
                                            codigo=str(row['Código Destino 1']),
                                            unidadesAFinanciar=str(row['Unidades Destino 1']),
                                            costoInversion=str(row['Costo Inversión 1']))
        
            # Crear el elemento 'destinoCreditoValorAFinanciar' dentro de 'destinoCredito'
            destino_credito_valor = ET.SubElement(destino_credito, "{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar")
            valor_a_financiar = ET.SubElement(destino_credito_valor, "{http://www.finagro.com.co/sit}valorAFinanciar", {"xmlns": ""})
            valor_a_financiar.text=str(row['Valor a Financiar 1'])
            
            if not row['Código Destino 2']!= row['Código Destino 2']: 
                # Crear un elemento 'destinoCredito' dentro de 'destinosCredito'
                destino_credito = ET.SubElement(destinos_credito, "{http://www.finagro.com.co/sit}destinoCredito",
                                                codigo=str(row['Código Destino 2']),
                                                unidadesAFinanciar=str(row['Unidades Destino 2']),
                                                costoInversion=str(row['Costo Inversión 2']))
        
                # Crear el elemento 'destinoCreditoValorAFinanciar' dentro de 'destinoCredito'
                destino_credito_valor = ET.SubElement(destino_credito, "{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar")
                valor_a_financiar = ET.SubElement(destino_credito_valor, "{http://www.finagro.com.co/sit}valorAFinanciar", {"xmlns": ""})
                valor_a_financiar.text=str(row['Valor a Financiar 2'])
            
            if not row['Código Destino 3']!= row['Código Destino 3']: 
                # Crear un elemento 'destinoCredito' dentro de 'destinosCredito'
                destino_credito = ET.SubElement(destinos_credito, "{http://www.finagro.com.co/sit}destinoCredito",
                                                codigo=str(row['Código Destino 3']),
                                                unidadesAFinanciar=str(row['Unidades Destino 3']),
                                                costoInversion=str(row['Costo Inversión 3']))
        
                # Crear el elemento 'destinoCreditoValorAFinanciar' dentro de 'destinoCredito'
                destino_credito_valor = ET.SubElement(destino_credito, "{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar")
                valor_a_financiar = ET.SubElement(destino_credito_valor, "{http://www.finagro.com.co/sit}valorAFinanciar", {"xmlns": ""})
                valor_a_financiar.text=str(row['Valor a Financiar 3'])
            
            if not row['Código Destino 4']!= row['Código Destino 3']: 
                # Crear un elemento 'destinoCredito' dentro de 'destinosCredito'
                destino_credito = ET.SubElement(destinos_credito, "{http://www.finagro.com.co/sit}destinoCredito",
                                                codigo=str(row['Código Destino 4']),
                                                unidadesAFinanciar=str(row['Unidades Destino 4']),
                                                costoInversion=str(row['Costo Inversión 4']))
        
                # Crear el elemento 'destinoCreditoValorAFinanciar' dentro de 'destinoCredito'
                destino_credito_valor = ET.SubElement(destino_credito, "{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar")
                valor_a_financiar = ET.SubElement(destino_credito_valor, "{http://www.finagro.com.co/sit}valorAFinanciar", {"xmlns": ""})
                valor_a_financiar.text=str(row['Valor a Financiar 4'])
            
            # Crear el elemento 'financiacion'
            financiacion = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}financiacion",
                                           fechaVencimientoFinal=str(fechaFinal),
                                           plazoCredito=str(row['Plazo']),
                                           valorTotalCredito=str(row['Capital Total']),
                                           porcentaje="100",
                                           valorObligacion=str(row['Capital Total']))
        
            # Datos de ejemplo para las cuotas
            datos_cuotas = []
            cantidad_cuotas= int(int(row['Plazo'])/int(row['Tipo Plan de Pagos']))
            cuota_capital = int(int(row['Capital Total'])/cantidad_cuotas)
            
            ult_cuota_capital = cuota_capital if int(row['Capital Total']) - (cantidad_cuotas*cuota_capital) == 0 else str(int(row['Capital Total']) - (Decimal(cantidad_cuotas-1)*Decimal(cuota_capital)))
            fHasta = pd.to_datetime(row['Fecha de Suscripción'],format ='%Y-%m-%d')
            for i in range(cantidad_cuotas-1):
                meses = int(row['Tipo Plan de Pagos'])
                fHasta = fHasta + relativedelta(months=meses)
                cuotas = {
                                "registro": str(i+1),
                                "fechaAplicacionHasta":str(date(int(fHasta.strftime('%Y')),int(fHasta.strftime('%m')),10)),
                                "conceptoRegistroCuota": "I" if tipo_plan == 1 else "K",
                                "periodicidadIntereses": "PE",
                                "periodicidadCapital": "" if tipo_plan == 1 else "PE",
                                "tasaBaseBeneficiario": "5",
                                "margenTasaBeneficiario": str(row['Puntos IBR']),
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
                        "margenTasaBeneficiario":  str(row['Puntos IBR']),
                        "valorCuotaCapital": str(row['Capital Total']) if tipo_plan == 1 else str(ult_cuota_capital),
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
        tree = ET.ElementTree(obligaciones)
        ET.indent(tree, space="  ", level=0)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp:
                    tree.write(tmp.name, encoding="UTF-8", xml_declaration=True)
                    st.success("✅ XML generado exitosamente.")
                    with open(tmp.name, "rb") as f:
                        st.download_button("📥 Descargar XML de Obligaciones Nuevas", f, file_name="Obligaciones.xml", mime="application/xml")
