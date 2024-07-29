import pandas as pd
import os
from modulos.process_docx import create_table_at_placeholder, replace_text, save_docx
from typing import Union
from modulos.process_pdf import encrypt_pdf
from docx2pdf import convert

def read_and_save_excel_to_list_df(input_folder: str)-> str:
    """
    Lee todos los archivos Excel (.xlsx y .xls) en una carpeta.

    Esta función busca en el directorio especificado todos los archivos con extensiones .xlsx o .xls, 
    los lee como DataFrames usando la biblioteca `pandas`.

    :param input_folder: Ruta al directorio que contiene los archivos Excel. Debe ser una ruta válida en el sistema de archivos.
    
    :return: Un DataFrame de pandas que contiene los datos combinados de todos los archivos Excel leídos.

    :raises FileNotFoundError: Si el directorio especificado no existe.
    :raises ValueError: Si no se encuentran archivos Excel en el directorio.
    """
    if not os.path.exists(input_folder):
        raise FileNotFoundError(f"La carpeta {input_folder} no existe.")
    
    # Lista para almacenar los DataFrames
    dataframes = []

    # Iterar sobre todos los archivos en la carpeta
    for file_name in os.listdir(input_folder):
        if file_name.endswith('.xlsx') or file_name.endswith('.xls'):  # Incluye .xls para archivos Excel antiguos
            file_path = os.path.join(input_folder, file_name)
            
            # Leer el archivo Excel y añadir el DataFrame a la lista
            df = pd.read_excel(file_path, engine='openpyxl')  # Usa 'xlrd' si es .xls
            dataframes.append(df)

    # Concatenar todos los DataFrames en uno solo
    all_data = pd.concat(dataframes, ignore_index=True)
    
    return all_data

def format_currency(value: Union[int, float]) -> str:
    """
    Formatea el valor como moneda en formato decimal con el signo '$'.
    Elimina los centavos si son cero.
    
    :param value: Valor a formatear.
    :return: Valor formateado como cadena con signo '$'.
    """
    if pd.isna(value) or value == 0:
        return "$0.00"
    
    formatted = "${:,.2f}".format(value)
    if formatted.endswith(".00"):
        return formatted[:-3]  # Elimina ".00"
    
    return formatted

def main():
    input_folder = 'ExcelToSecurePDF/input'
    data = read_and_save_excel_to_list_df(input_folder)

    # Separamos el df por rut
    ruts = data["RUT"].unique()

    for rut in ruts:
        # Filtramos el df por rut
        df = data[data["RUT"] == rut]

        # Verificar si df está vacío
        if df.empty:
            print(f"No hay datos para el RUT {rut}.")
            continue  # Salta al siguiente RUT

        # Verificar la existencia de columnas
        required_columns = ["RAZON_SOCIAL_PC", "FEC_CAL_PC", "MONTO_NOMINAL_PC", "INTERES_PC", "MONTO_ACT_PC"]
        for col in required_columns:
            if col not in df.columns:
                raise KeyError(f"Falta la columna {col} en el DataFrame para el RUT {rut}.")

        # Manejo de datos nulos (opcional)
        df = df.fillna({
            "RAZON_SOCIAL_PC": "",
            "FEC_CAL_PC": "",
            "MONTO_NOMINAL_PC": 0,
            "INTERES_PC": 0,
            "MONTO_ACT_PC": 0
        })

        # Filtrar filas donde los valores en las columnas relevantes no sean todos 0
        df = df[~((df["MONTO_NOMINAL_PC"] == 0) & (df["INTERES_PC"] == 0) & (df["MONTO_ACT_PC"] == 0))]

        # Verificar si df está vacío después del filtrado
        if df.empty:
            print(f"Después del filtrado, no hay datos relevantes para el RUT {rut}.")
            continue  # Salta al siguiente RUT

        # Verificar si hay al menos una fila en df antes de acceder a los valores
        if not df["RAZON_SOCIAL_PC"].empty:
            replacements = {
                "{RAZON_SOCIAL_PC}": df["RAZON_SOCIAL_PC"].values[0],
                "{RUT}": df["RUT"].values[0]
            }
        else:
            print(f"No se puede encontrar 'RAZON_SOCIAL_PC' para el RUT {rut}.")
            continue  # Salta al siguiente RUT

        # Extraer datos y convertir a listas con formato de moneda
        data_fecha_pago = df["FEC_CAL_PC"].astype(str).tolist()
        data_monto_nominal = df["MONTO_NOMINAL_PC"].apply(format_currency).tolist()
        data_intereses_reajustes = df["INTERES_PC"].apply(format_currency).tolist()
        data_monto_actualizado = df["MONTO_ACT_PC"].apply(format_currency).tolist()

        # Construir la tabla de datos para el .docx
        table_data = [
            ["Fecha Pago", "Monto nominal", "Intereses y reajustes", "Monto actualizado"]
        ]
        
        for i in range(len(data_fecha_pago)):
            table_data.append([
                str(data_fecha_pago[i]),
                str(data_monto_nominal[i]),
                str(data_intereses_reajustes[i]),
                str(data_monto_actualizado[i])
            ])

        # Reemplazamos los valores en el archivo .docx
        doc = replace_text("ExcelToSecurePDF/template/doc_template.docx", replacements)

        # Creamos la tabla con los datos del df
        doc = create_table_at_placeholder(doc, table_data, "{tabla}")

        # Guardamos el archivo .docx
        save_docx(doc, f"ExcelToSecurePDF/output/{rut}.docx")

        # Convertimos el .docx a .pdf
        convert(f"ExcelToSecurePDF/output/{rut}.docx", f"ExcelToSecurePDF/output/{rut}.pdf")

        # Borramos el .docx
        os.remove(f"ExcelToSecurePDF/output/{rut}.docx")

        password = rut.split("-")[0].replace(".", "")

        # Encryptamos el pdf
        encrypt_pdf(f"ExcelToSecurePDF/output/{rut}.pdf", password)

if __name__ == "__main__":
    main()
