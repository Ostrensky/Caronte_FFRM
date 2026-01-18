# --- FILE: app/pgdas_loader.py ---

import re
import glob
import os
import logging
from PyPDF2 import PdfReader

def _read_pdf_text(file_path: str) -> str:
    # ... (this function remains the same)
    try:
        text = ""
        with open(file_path, "rb") as pdf_file:
            reader = PdfReader(pdf_file)
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text
        return text
    except Exception as e:
        logging.warning(f"Não foi possível ler o PDF '{file_path}': {e}")
        return ""

def _extract_iss_value(document_text: str) -> float | None:
    # ... (this function remains the same)
    pattern = re.compile(
        r'Total do Débito Exigível \(R\$\)' # Target header
        r'[\s\S]*?'                       # Non-greedy match
        r'(\d[.,\d]*\d)\s+'               # 1. IRPJ
        r'(\d[.,\d]*\d)\s+'               # 2. CSLL
        r'(\d[.,\d]*\d)\s+'               # 3. COFINS
        r'(\d[.,\d]*\d)\s+'               # 4. PIS/Pasep
        r'(\d[.,\d]*\d)\s+'               # 5. INSS/CPP
        r'(\d[.,\d]*\d)\s+'               # 6. ICMS
        r'(\d[.,\d]*\d)\s+'               # 7. IPI
        r'(\d[.,\d]*\d)\s+'               # 8. ISS (the value we want)
        r'(\d[.,\d]*\d)'                  # 9. Total
    )
    match = pattern.search(document_text)
    if match:
        iss_string = match.group(8)
        iss_value = float(iss_string.replace('.', '').replace(',', '.'))
        return iss_value
    return None

def _extract_pa_date(document_text: str) -> str | None:
    # ... (this function remains the same)
    pattern_1 = re.compile(
        r"Período de Apuração" r"\s*(?:\(PA\))?" r"\s*:\s*"
        r"\d{2}\s*/\s*" r"(\d{2})" r"\s*/\s*" r"(\d{4})"
    )
    match_1 = pattern_1.search(document_text)
    if match_1:
        month = match_1.group(1); year = match_1.group(2)
        return f"{month}/{year}"
    pattern_2 = re.compile(r"Período de Apuração \(PA\)\s*(\d{2}/\d{4})")
    match_2 = pattern_2.search(document_text)
    if match_2:
        return match_2.group(1)
    return None

# ✅ --- START: New function to extract Declaration Number ---
def _extract_declaration_number(document_text: str) -> str | None:
    """Extracts the 'Nº da Declaração'."""
    # Pattern looks for "Nº da Declaração:" followed by digits
    pattern = re.compile(r"Nº da Declaração:\s*(\d+)")
    match = pattern.search(document_text)
    if match:
        # Return the captured group (the number)
        return match.group(1)
    return None
# ✅ --- END: New function ---

def _load_and_process_pgdas(folder_path, status_callback=None):
    """
    Reads all PGDASD PDFs, extracts ISS, PA date, and Declaration Number.
    Returns a dictionary mapping 'MM/YYYY' to a tuple: (total_iss_payment, declaration_number).
    """
    emit = status_callback.emit if status_callback else print

    if not folder_path or not os.path.isdir(folder_path):
        return {}

    pdf_files = glob.glob(os.path.join(folder_path, 'PGDASD*.pdf'))
    if not pdf_files:
        emit("Nenhum ficheiro PGDASD*.pdf encontrado na pasta selecionada.")
        return {}

    # ✅ Map stores: MM/YYYY -> (total_amount, declaration_number)
    pgdas_payments_map = {}
    emit(f"A processar {len(pdf_files)} ficheiros PGDAS...")

    for pdf_path in pdf_files:
        filename = os.path.basename(pdf_path)
        try:
            document_text = _read_pdf_text(pdf_path)
            if not document_text:
                continue

            iss_value = _extract_iss_value(document_text)
            pa_date = _extract_pa_date(document_text)
            declaration_number = _extract_declaration_number(document_text) # ✅ Extract number

            if iss_value is not None and pa_date:
                # Get current data or defaults
                current_amount, current_decl_num = pgdas_payments_map.get(pa_date, (0.0, "-"))

                # Sum the amount
                new_amount = current_amount + iss_value

                # Keep the first valid declaration number found for the month
                new_decl_num = current_decl_num
                if current_decl_num == "-" and declaration_number:
                    new_decl_num = declaration_number

                pgdas_payments_map[pa_date] = (new_amount, new_decl_num) # ✅ Store tuple

            # Update logging messages
            elif not pa_date:
                emit(f"  - Aviso: 'Período de Apuração' não encontrado em {filename}")
            elif iss_value is None: # Changed condition slightly
                 emit(f"  - Aviso: Valor de ISS não encontrado em {filename} (PA: {pa_date})")
            # Log if declaration number is missing (optional)
            # elif not declaration_number:
            #     emit(f"  - Info: 'Nº da Declaração' não encontrado em {filename} (PA: {pa_date})")


        except Exception as e:
            emit(f"  - Erro ao processar {filename}: {e}")

    emit(f"Processamento PGDAS concluído. {len(pgdas_payments_map)} meses com pagamentos encontrados.")
    return pgdas_payments_map # Returns dict{ MM/YYYY -> (amount, decl_num) }