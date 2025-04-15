# -*- coding: utf-8 -*-
"""
Created on Tue Apr 15 07:39:38 2025

@author: Miguel Rico
"""

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from datetime import datetime
import numpy as np
import io
import os
import tempfile

def generate_report(df, author_name="Oscar Ivan Vargas Pineda"):
    """Generate a PDF report from the provided DataFrame"""
    
    # Create a temporary file for the PDF
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
        pdf_path = tmp_pdf.name
    
    # Create a temporary file for the chart
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_chart:
        chart_path = tmp_chart.name
    
    # Get report information
    total_inscritos = df['Nombre y apellidos completos'].nunique()
    
    # Convert date column to datetime if it's not already
    if df['Hora de inicio'].dtype != 'datetime64[ns]':
        df['Hora de inicio'] = pd.to_datetime(df['Hora de inicio'])
    
    fecha_min = df['Hora de inicio'].min().strftime('%d/%m/%Y %H:%M')
    fecha_max = df['Hora de inicio'].max().strftime('%d/%m/%Y %H:%M')
    
    # Count enrollments by course
    conteo_cursos = df.groupby('Curso de interés').agg({'Nombre y apellidos completos': 'nunique'})
    conteo_cursos = conteo_cursos.rename(columns={'Nombre y apellidos completos': 'Cantidad de inscritos'})
    conteo_cursos = conteo_cursos.sort_values('Cantidad de inscritos', ascending=False)
    
    # Create bar chart
    plt.figure(figsize=(12, 8))
    bars = plt.bar(conteo_cursos.index, conteo_cursos['Cantidad de inscritos'], color='steelblue')
    
    # Add labels to bars
    for bar in bars:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                 f'{height:.0f}', ha='center', va='bottom', fontweight='bold')
    
    plt.xticks(rotation=45, ha='right', fontsize=10)
    plt.xlabel('Curso', fontsize=12)
    plt.ylabel('Cantidad de inscritos', fontsize=12)
    plt.title('Número de inscritos por curso', fontsize=14, fontweight='bold')
    plt.tight_layout()
    
    # Save chart temporarily
    plt.savefig(chart_path, dpi=150, bbox_inches='tight')
    plt.close()
    
    # Create PDF with ReportLab
    doc = SimpleDocTemplate(pdf_path, pagesize=letter)
    styles = getSampleStyleSheet()
    
    # Create custom styles
    styles.add(ParagraphStyle(name='TituloReporte', fontSize=18, alignment=1, spaceAfter=12, fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle(name='SubtituloReporte', fontSize=14, spaceAfter=8, fontName="Helvetica-Bold"))
    styles['Normal'].fontSize = 11
    styles['Normal'].spaceAfter = 6
    styles.add(ParagraphStyle(name='FechaReporte', fontSize=10, alignment=2, textColor=colors.gray))
    styles.add(ParagraphStyle(name='AutorReporte', fontSize=10, alignment=2, textColor=colors.gray))
    
    # PDF elements
    elementos = []
    
    # Title and date
    fecha_actual = datetime.now().strftime("%d/%m/%Y")
    elementos.append(Paragraph("REPORTE DE INSCRIPCIONES A CURSOS", styles['TituloReporte']))
    elementos.append(Paragraph(f"Fecha de elaboración: {fecha_actual}", styles['FechaReporte']))
    elementos.append(Paragraph(f"Elaborado por: {author_name}", styles['AutorReporte']))
    elementos.append(Spacer(1, 0.25 * inch))
    
    # General information
    elementos.append(Paragraph("Información General", styles['SubtituloReporte']))
    elementos.append(Paragraph(f"Total de personas inscritas: <b>{total_inscritos}</b>", styles['Normal']))
    elementos.append(Paragraph(f"Fecha de inicio de inscripciones: <b>{fecha_min}</b>", styles['Normal']))
    elementos.append(Paragraph(f"Fecha de finalización de inscripciones: <b>{fecha_max}</b>", styles['Normal']))
    elementos.append(Spacer(1, 0.25 * inch))
    
    # Bar chart
    elementos.append(Paragraph("Distribución de Inscripciones por Curso", styles['SubtituloReporte']))
    elementos.append(Image(chart_path, width=7*inch, height=4.5*inch))
    elementos.append(Spacer(1, 0.25 * inch))
    
    # Tables for each course
    elementos.append(Paragraph("Detalle de Inscritos por Curso", styles['SubtituloReporte']))
    
    for curso in conteo_cursos.index:
        # Filter data for this course
        curso_df = df[df['Curso de interés'] == curso].copy()
        
        # Remove duplicates
        curso_df = curso_df.drop_duplicates(subset=['Nombre y apellidos completos', 'Correo de contacto'])
        
        # Sort by name
        curso_df = curso_df.sort_values('Nombre y apellidos completos')
        
        # Get table data
        datos_tabla = [["Nombre y Apellidos", "Correo de contacto"]]
        
        for _, row in curso_df.iterrows():
            datos_tabla.append([row['Nombre y apellidos completos'], row['Correo de contacto']])
        
        # Add course title
        elementos.append(Paragraph(f"Curso: {curso} ({len(datos_tabla) - 1} inscritos)", styles['Normal']))
        
        # Create table
        tabla = Table(datos_tabla, colWidths=[3.5*inch, 3.5*inch])
        tabla.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
        ]))
        
        elementos.append(tabla)
        elementos.append(Spacer(1, 0.25 * inch))
    
    # Create PDF
    doc.build(elementos)
    
    # Clean up chart file after building PDF
    os.unlink(chart_path)
    
    return pdf_path


def main():
    st.set_page_config(page_title="Generador de Reportes de Inscripciones", page_icon="📊", layout="wide")
    
    st.title("📊 Generador de Reportes de Inscripciones a Cursos")
    st.markdown("""
    Esta aplicación genera un reporte PDF detallado de inscripciones a cursos a partir de un archivo Excel.
    
    Simplemente sube tu archivo Excel con los datos de inscripción y haz clic en 'Generar Reporte'.
    """)
    
    # File uploader
    uploaded_file = st.file_uploader("Carga tu archivo Excel de inscripciones", type=["xlsx"])
    
    # Author name input
    author_name = st.text_input("Nombre del autor del reporte", "Oscar Ivan Vargas Pineda")
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            
            # Check if the file has the required columns
            required_columns = ['Nombre y apellidos completos', 'Correo de contacto', 'Curso de interés', 'Hora de inicio']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"El archivo Excel no contiene las siguientes columnas requeridas: {', '.join(missing_columns)}")
            else:
                st.success("Archivo cargado correctamente")
                
                # Show data preview
                st.subheader("Vista previa de los datos")
                st.dataframe(df.head())
                
                # General statistics
                st.subheader("Estadísticas generales")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Inscritos", df['Nombre y apellidos completos'].nunique())
                with col2: 
                    st.metric("Total Cursos", df['Curso de interés'].nunique())
                with col3:
                    if 'Hora de inicio' in df.columns:
                        if df['Hora de inicio'].dtype != 'datetime64[ns]':
                            df['Hora de inicio'] = pd.to_datetime(df['Hora de inicio'])
                        date_range = f"{(df['Hora de inicio'].max() - df['Hora de inicio'].min()).days} días"
                        st.metric("Rango de Fechas", date_range)
                
                # Generate report button
                if st.button("Generar Reporte PDF"):
                    with st.spinner("Generando reporte..."):
                        pdf_path = generate_report(df, author_name)
                        
                        # Read the generated PDF
                        with open(pdf_path, "rb") as pdf_file:
                            pdf_bytes = pdf_file.read()
                        
                        # Clean up the temporary PDF file
                        os.unlink(pdf_path)
                        
                        # Create download button
                        st.download_button(
                            label="Descargar Reporte PDF",
                            data=pdf_bytes,
                            file_name="reporte_inscritos.pdf",
                            mime="application/pdf"
                        )
                        
                        st.success("¡Reporte generado con éxito!")
        
        except Exception as e:
            st.error(f"Error al procesar el archivo: {str(e)}")
            st.exception(e)
    
    # Add information section
    st.sidebar.header("Información")
    st.sidebar.info("""
    **Formato del archivo Excel requerido**
    
    El archivo Excel debe contener las siguientes columnas:
    - Nombre y apellidos completos
    - Correo de contacto
    - Curso de interés
    - Hora de inicio
    
    Las fechas deben estar en un formato reconocible por pandas.
    """)
    
    st.sidebar.header("Acerca de")
    st.sidebar.markdown("""
    Esta aplicación fue creada como una herramienta para generar reportes
    de inscripciones a cursos a partir de datos en Excel.
    
    Desarrollado con Streamlit y ReportLab.
    """)

if __name__ == "__main__":
    main()