from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

GREEN = "92d050"
FONT_HEADER = "Calibri"
FONT_HEADER_SIZE = 11
FONT_CONTENT = "Calibri"
FONT_CONTENT_SIZE = 11

HEADER_FOR_REQUESTS_PAGES = [
    "Solicitud",
    "Convocatoria Bolsa Concursal",
    "N. Comité",
    "Radicado de la solicitud",
    "Docente",
    "Programa",
    "Facultad",
    "Evento",
    "Fecha viaje",
    "Tipo de movilidad",
    "Destino",
    "Institución de destino",
    "Descripción del evento y la participación",
    "Monto solicitado",
    "Rubros",
    "Compromisos",
    "Observaciones",
    "Concepto del Comité de Movilidad",
    "Respuesta radicado",
    "Final solicitado en Abedul por la Facultad",
    "Evidencia Abedul",
    "Acta de compromisos",
    "Observaciones de la solicitud"
]

HEADER_FOR_LICENSES_PAGES = [
    "Solicitud",
    "Convocatoria Bolsa Concursal",
    "N. Comité",
    "Radicado de la solicitud",
    "Docente",
    "Programa",
    "Facultad",
    "Evento",
    "Fecha viaje",
    "Tipo de movilidad",
    "Destino",
    "Institución de destino",
    "Descripción del evento y la participación",
    "Monto solicitado",
    "Rubros",
    "Compromisos",
    "Observaciones",
    "Concepto del Comité de Movilidad",
    "Respuesta radicado",
    "Final solicitado en Abedul por la Facultad",
    "Evidencia Abedul",
    "Acta de compromisos",
    "Observaciones de la solicitud"
]

CONSOLIDATED_HEADER = [
    "Docente",
    "Monto Solicitado",
    "Facultad"
]

def create_header_style():
    """
    Creating the Style of the Excel Header.
    """

    font = Font(name = FONT_HEADER, 
                size = FONT_HEADER_SIZE,
                bold = True)
    fill = PatternFill(start_color = GREEN,
                end_color = GREEN,
                fill_type = "solid")
    alignment = Alignment(horizontal = "center", 
                vertical = "center", 
                wrap_text = True)
    border = Border(left = Side(border_style = "thin", color = "000000"),
                right = Side(border_style = "thin", color = "000000"),
                top = Side(border_style = "thin", color = "000000"),
                bottom = Side(border_style = "thin", color = "000000"))

    return font, fill, alignment, border

def create_content_style():
    """
    Creating the Style of the Excel Content.
    """
    
    font = Font(name = FONT_CONTENT,
                size = FONT_CONTENT_SIZE,
                bold = False)
    alignment = Alignment(horizontal = "left",
                vertical = "center",
                wrap_text = True)
    border = Border(left = Side(border_style = "thin", color = "000000"),
                right = Side(border_style = "thin", color = "000000"),
                top = Side(border_style = "thin", color = "000000"),
                bottom = Side(border_style = "thin", color = "000000"))
    
    return font, alignment, border

def apply_style_to_header(cell):
    """
    Apply the Style to the Excel Header.
    """

    font, fill, alignment, border = create_header_style()

    cell.font = font
    cell.fill = fill
    cell.alignment = alignment
    cell.border = border

def apply_style_to_content(cell):
    """
    Apply the Style to the Excel Content.
    """

    font, alignment, border = create_content_style()

    cell.font = font
    cell.alignment = alignment
    cell.border = border

def adjust_column_width(worksheet, column = None):
    """
    Adjust the width of the columns in the Excel.
    """

    if column:
        for i, ancho in enumerate(column, 1):
            letter = get_column_letter(i)
            worksheet.column_dimensions[letter].width = ancho
    else:
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass

                max_length = min(max_length, 57)
                worksheet.column_dimensions[column_letter].width = max_length + 5