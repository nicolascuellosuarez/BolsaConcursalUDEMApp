"""
Module to extract data from Bolsa Concursal PDFs.
"""

import re
import pdfplumber
from typing import Dict, Optional, List


def extract_pdf_text(pdf_path: str) -> str:
    """Extract all text from a PDF using pdfplumber."""
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
    return full_text


def extract_pdf_text_by_page(pdf_path: str) -> List[str]:
    """Extract text from each page individually."""
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            pages_text.append(text if text else "")
    return pages_text


def extract_radicado(text: str, pages_text: Optional[List[str]] = None) -> str:
    """
    Extract radicado number from the header of the first page.
    The header usually contains: "C.I Radicado: [number] Fecha: [date]"
    """
    if pages_text and len(pages_text) > 0:
        first_page = pages_text[0]
    else:
        first_page = text
    
    pattern = r"C\.?I\.?\s*Radicado:?\s*([\d\-/]+)"
    match = re.search(pattern, first_page, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else ""


def extract_docente(text: str) -> str:
    """Extract docente name."""
    pattern = r"Mi nombre es (.*?)(?:,| docente|\.)"
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else ""


def extract_facultad(text: str) -> str:
    """Extract faculty name."""
    pattern = r"Facultad de ([A-Za-zÁÉÍÓÚáéíóúÑñ\s]+?)(?=\.|,|\n| y|$)"
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        result = match.group(1).strip()
        if len(result.split()) > 3:
            first_words = result.split()[:2]
            return " ".join(first_words)
        return result
    return ""

def extract_evento(text: str) -> str:
    """Extract event name."""
    pattern = r"Nombre del evento:?\s*(.*?)(?=\n|Universidad|$)"
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else ""


def extract_fechas_evento(text: str) -> str:
    """Extract event dates."""
    pattern = r"Fechas del evento:?\s*(.*?)(?=\n|Descripción|$)"
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else ""


def extract_institucion(text: str) -> str:
    """Extract host institution."""
    patterns = [
        r"Universidad/institución anfitriona:?\s*(.*?)(?=\n|Ciudad|$)",
        r"Institución anfitriona:?\s*(.*?)(?=\n|Ciudad|$)",  # ← NUEVO: sin "Universidad/"
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return ""


def extract_destino(text: str) -> str:
    """Extract city and country."""
    pattern = r"Ciudad, país:?\s*(.*?)(?=\n|Fechas|$)"
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else ""


def extract_tipo_movilidad(destino: str) -> str:
    """Determine mobility type based on destination."""
    if not destino:
        return ""
    return "Nacional" if "colombia" in destino.lower() else "Internacional"


def extract_descripcion(text: str) -> str:
    """Extract event description and participation."""
    pattern1 = r"Descripción general del evento:?\s*(.*?)(?=Descripción de la participación|En el marco|Para llevar a cabo|$)"
    match1 = re.search(pattern1, text, re.IGNORECASE | re.DOTALL)
    desc_general = match1.group(1).strip() if match1 else ""
    
    pattern2 = r"Descripción de la participación y actividades a desarrollar:?\s*(.*?)(?=Para llevar a cabo|En el marco|De igual manera|$)"
    match2 = re.search(pattern2, text, re.IGNORECASE | re.DOTALL)
    desc_participacion = match2.group(1).strip() if match2 else ""
    
    if desc_general and desc_participacion:
        return f"{desc_general}\n\n{desc_participacion}"
    return desc_general or desc_participacion


def extract_monto_total(text: str) -> str:
    """Extract total amount requested."""
    pattern = r"Total de la solicitud:?\s*([\d\.\,]+)"
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else ""


def extract_rubros(text: str) -> str:
    """Extract cost items from the table."""
    patterns = [
        r"Concepto.*?Total de la solicitud.*?(?:\n|$)",  # Original
        r"Concepto.*?Total solicitud.*?(?:\n|$)",       # ← NUEVO: sin "de la"
        r"Concepto.*?Total de la solicitud.*?(?:\n|$)",  # Fallback
    ]
    
    table_text = ""
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            table_text = match.group(0)
            break
    
    if not table_text:
        return ""
    
    lines = table_text.split('\n')
    rubros = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # Saltar encabezados (más flexible)
        if any(keyword in line for keyword in ["Concepto", "Valor", "Cantidad", "Total", "Valor estimado"]):
            continue
        if "Total de la solicitud" in line or "Total solicitud" in line:
            continue
        if "no aplica" in line.lower():
            continue
        if len(line) > 5:
            rubro = re.sub(r'\s+', ' ', line).strip()
            if rubro and not rubro.startswith('$'):
                rubros.append(rubro)
    
    return "\n".join(rubros) if rubros else ""


def extract_compromisos(text: str) -> str:
    """
    Extract ONLY the commitments related to the university.
    """
    pattern_section = r"(?:En el marco de (?:mi|esta) participación (?:en dicho evento)?,?\s*)?mis\s+compromisos\s+con\s+la\s+Universidad\s+de\s+Medellín\s+serán\s+los\s+siguientes:?\s*(.*?)(?=Agradezco|De igual manera|Finalmente|Para llevar a cabo|$)"
    match = re.search(pattern_section, text, re.IGNORECASE | re.DOTALL)
    
    if not match:
        pattern_section2 = r"mis\s+compromisos\s+con\s+la\s+Universidad\s+de\s+Medellín\s+serán\s+los\s+siguientes:?\s*(.*?)(?=Agradezco|De igual manera|Para llevar a cabo|$)"
        match = re.search(pattern_section2, text, re.IGNORECASE | re.DOTALL)
    
    if not match:
        return ""
    
    section_text = match.group(1).strip()
    
    pattern_items = r"(\d+\.\s.*?)(?=\n\d+\.|\n[A-ZÁÉÍÓÚ][a-záéíóú]|$)"
    items = re.findall(pattern_items, section_text, re.DOTALL)
    
    if items:
        cleaned_items = []
        for item in items:
            clean = re.sub(r'\s+', ' ', item.strip())
            cleaned_items.append(clean)
        return "\n".join(cleaned_items)
    
    return ""


def extract_pdf_data(pdf_path: str) -> Dict[str, str]:
    """
    Extract all fields from PDF and return as dictionary.
    """
    text = extract_pdf_text(pdf_path)
    pages_text = extract_pdf_text_by_page(pdf_path)
    
    if not text:
        return {}
    
    docente = extract_docente(text)
    facultad = extract_facultad(text)
    destino = extract_destino(text)
    
    return {
        "radicado": extract_radicado(text, pages_text),
        "docente": docente,
        "facultad": facultad,
        "evento": extract_evento(text),
        "fechas_evento": extract_fechas_evento(text),
        "institucion": extract_institucion(text),
        "destino": destino,
        "tipo_movilidad": extract_tipo_movilidad(destino),
        "descripcion": extract_descripcion(text),
        "monto_total": extract_monto_total(text),
        "rubros": extract_rubros(text),
        "compromisos": extract_compromisos(text),
    }


def print_data(data: Dict[str, str]):
    """Print extracted data in a readable format."""
    print("\n" + "="*60)
    print("Información Extraída del PDF.")
    print("="*60)
    for key, value in data.items():
        print(f"{key}: {value}")
    print("="*60)



if __name__ == "__main__":
    pdf_path = "../data/Ana Milena Montoya Ruiz.pdf" 
    print("Extrayendo Información del PDF.")
    data = extract_pdf_data(pdf_path)
    print_data(data)