import pandas as pd

def process_xlsx(file_path: str) -> list:
    """Extract the text from a .xlsx file."""
    # Extraemos la información de la hoja de cálculo y su encabezado
    df = pd.read_excel(file_path)
    return df
