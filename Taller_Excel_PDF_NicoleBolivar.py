"""Taller segundo momento - Nicole Bolivar - TecnoVentas."""
from pathlib import Path
from html import escape
from math import isfinite

import openpyxl
from tabulate import tabulate
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

CARPETA = Path(__file__).resolve().parent
EXCEL = CARPETA / 'ventas_tecnologia_v2.xlsx'
PDF = CARPETA / 'reporte_ejecutivo_ventas.pdf'
ENCABEZADOS = ['ID', 'Producto', 'Categoria', 'Precio ($)', 'Ventas (Unid)']
ALINEACION = ('center', 'left', 'center', 'right', 'right')


def cargar_datos(ruta=EXCEL):
    """Lee resultados guardados de Excel y recorre las filas dinamicamente."""
    libro = openpyxl.load_workbook(ruta, data_only=True)
    datos = []
    try:
        hoja = libro.active
        for numero, fila in enumerate(hoja.iter_rows(min_row=2, values_only=True), 2):
            if all(valor is None for valor in fila):
                continue
            if len(fila) < 5 or any(valor is None for valor in fila[:5]):
                raise ValueError(f'Fila {numero}: faltan datos. Si hay formulas, recalcula y guarda en Excel.')
            registro = list(fila[:5])
            identificador, producto, categoria, precio, ventas = registro
            if isinstance(identificador, bool) or int(identificador) != float(identificador):
                raise ValueError(f'Fila {numero}: el ID debe ser entero.')
            if not str(producto).strip() or not str(categoria).strip():
                raise ValueError(f'Fila {numero}: producto o categoria vacios.')
            if isinstance(precio, bool) or isinstance(ventas, bool):
                raise ValueError(f'Fila {numero}: precio o ventas invalidos.')
            precio, ventas = float(precio), float(ventas)
            if not isfinite(precio) or not isfinite(ventas) or precio < 0 or ventas < 0 or not ventas.is_integer():
                raise ValueError(f'Fila {numero}: precio y unidades deben ser numeros no negativos; unidades enteras.')
            registro = [int(identificador), str(producto), str(categoria), precio, int(ventas)]
            if any(anterior[0] == registro[0] for anterior in datos):
                raise ValueError(f'Fila {numero}: ID duplicado.')
            datos.append(registro)
    finally:
        libro.close()
    return datos


def filtrar_precio(datos):
    return [fila for fila in datos if fila[3] > 100]


def calcular_metricas(datos):
    unidades = sum(fila[4] for fila in datos)
    promedio = sum(fila[3] for fila in datos) / len(datos) if datos else 0
    return unidades, promedio


def mostrar_tabla(datos, titulo='Catalogo completo'):
    print('\n' + titulo)
    if datos:
        filas = [[f[0], f[1], f[2], f'${f[3]:,.2f}', f[4]] for f in datos]
        print(tabulate(filas, headers=ENCABEZADOS, tablefmt='fancy_grid', colalign=ALINEACION))
    else:
        print('No hay productos para mostrar.')
    unidades, promedio = calcular_metricas(datos)
    print(f'Productos: {len(datos)} | Total unidades vendidas: {unidades} | Precio promedio: ${promedio:,.2f}')


def buscar_por_id(datos):
    try:
        identificador = int(input('Ingresa el ID del producto: '))
    except ValueError:
        print('Advertencia: debes ingresar un ID entero.')
        return
    resultado = [fila for fila in datos if fila[0] == identificador]
    if resultado:
        mostrar_tabla(resultado, f'Producto con ID {identificador}')
    else:
        print(f'Advertencia: no existe un producto con ID {identificador}.')


def exportar_pdf(datos, ruta=PDF, titulo='Catalogo completo'):
    estilos = getSampleStyleSheet()
    cuerpo = ParagraphStyle('Celda', parent=estilos['Normal'], fontSize=9, leading=12)
    centro = ParagraphStyle('Centro', parent=cuerpo, alignment=1)
    derecha = ParagraphStyle('Derecha', parent=cuerpo, alignment=2)
    encabezado = ParagraphStyle('Cabecera', parent=centro, textColor=colors.white, fontName='Helvetica-Bold')
    doc = SimpleDocTemplate(str(ruta), pagesize=letter, rightMargin=72, leftMargin=72,
                            topMargin=54, bottomMargin=54)
    story = [Paragraph('Informe Ejecutivo de Ventas TecnoVentas', estilos['Title']),
             Paragraph('Nicole Bolivar | Nuevas Tecnologias', estilos['Normal']), Spacer(1, 12),
             Paragraph(escape(titulo), estilos['Heading2']),
             Paragraph(f'Total de productos procesados: {len(datos)}', estilos['Normal']),
             Paragraph('Origen: ventas_tecnologia_v2.xlsx. La base incluida es de ejemplo; '
                       'reemplazarla por el archivo del docente para la entrega definitiva.', estilos['Normal']),
             Spacer(1, 12)]
    filas = [[Paragraph(escape(h), encabezado) for h in ENCABEZADOS]]
    for fila in datos:
        valores = [str(fila[0]), fila[1], fila[2], f'${fila[3]:,.2f}', str(fila[4])]
        filas.append([Paragraph(escape(valor), estilo) for valor, estilo in
                      zip(valores, (centro, cuerpo, centro, derecha, derecha))])
    unidades, promedio = calcular_metricas(datos)
    resumen = ['Total', f'{len(datos)} productos', 'Promedio / Total', f'${promedio:,.2f}', str(unidades)]
    filas.append([Paragraph(escape(valor), estilo) for valor, estilo in
                  zip(resumen, (centro, cuerpo, centro, derecha, derecha))])
    tabla = Table(filas, colWidths=[40, 160, 100, 80, 80], repeatRows=1)
    tabla.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1A365D')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), .5, colors.HexColor('#CBD5E0')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#F7FAFC')]),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#E2E8F0')),
        ('TOPPADDING', (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 7),
    ]))
    story.append(tabla)
    story.append(Spacer(1, 12))
    story.append(Paragraph('El precio promedio es la media simple de los precios unitarios '
                           'de los productos mostrados. La ultima columna del resumen suma las unidades vendidas.', estilos['Normal']))
    doc.build(story)
    print(f'PDF generado correctamente: {ruta.name}')


def menu(datos):
    # Menu interactivo del reto 4.
    while True:
        print('\nTECNO VENTAS\n1) Ver catalogo completo\n2) Filtrar por precio (> $100)'
              '\n3) Buscar por ID\n4) Exportar informe PDF\n5) Salir del programa')
        opcion = input('Selecciona una opcion: ').strip()
        if opcion == '1':
            mostrar_tabla(datos)
        elif opcion == '2':
            filtrados = filtrar_precio(datos)
            mostrar_tabla(filtrados, 'Productos con precio superior a $100')
            exportar_pdf(filtrados, CARPETA / 'reporte_ventas_filtrado.pdf', 'Productos con precio superior a $100')
        elif opcion == '3':
            buscar_por_id(datos)
        elif opcion == '4':
            exportar_pdf(datos)
        elif opcion == '5':
            print('Programa finalizado.')
            break
        else:
            print('Opcion invalida. Selecciona un numero del 1 al 5.')


def main():
    try:
        datos = cargar_datos()
    except (OSError, ValueError, TypeError, OverflowError, openpyxl.utils.exceptions.InvalidFileException) as error:
        print(f'No se pudo cargar el Excel: {error}')
        return
    # Flujo base del tutorial: Excel -> consola -> PDF automaticamente.
    mostrar_tabla(datos)
    exportar_pdf(datos)
    menu(datos)


if __name__ == '__main__':
    try:
        main()
    except (EOFError, KeyboardInterrupt):
        print('\nPrograma finalizado.')
