from modulos.process_docx import create_table_at_placeholder, replace_text, save_docx
from modulos.process_xlsx import process_xlsx
from modulos.process_pdf import encrypt_pdf
from docx2pdf import convert
import os

def main():
    # Procesamos el archivo .xlsx
    data = process_xlsx("input/comunicacion_mayo.xlsx")

    # Separamos el df por rut
    ruts = data["RUT"].unique()

    for rut in ruts:
        # Filtramos el df por rut
        df = data[data["RUT"] == rut]

        # Creamos el diccionario con los valores a reemplazar
        replacements = {
            "{RAZON_SOCIAL_PC}": df["RAZON_SOCIAL_PC"].values[0],
            "{RUT}": df["RUT"].values[0]
        }

        data_fecha_pago = df["FEC_CAL_PC"].tolist()
        data_monto_nominal = df["MONTO_NOMINAL_PC"].tolist()
        data_intereses_reajustes = df["INTERES_PC"].tolist()
        data_monto_actualizado = df["MONTO_ACT_PC"].tolist()

        table_data = [
            ["Fecha Pago", "Monto nominal", "Intereses y reajustes", "Monto actualizado"]
            ]
        
        for i in range(len(data_fecha_pago)):
            table_data.append([str(data_fecha_pago[i]),
                               str(data_monto_nominal[i]),
                               str(data_intereses_reajustes[i]),
                               str(data_monto_actualizado[i])])

        # Reemplazamos los valores en el archivo .docx
        doc = replace_text("template/Doc_Tipo.docx", replacements)

        # Creamos la tabla con los datos del df
        doc = create_table_at_placeholder(doc, table_data, "{tabla}")

        # Guardamos el archivo .docx
        save_docx(doc, f"output/{rut}.docx")

        # Convertimos el .docx a .pdf
        convert(f"output/{rut}.docx", f"output/{rut}.pdf")

        # Borramos el .docx
        os.remove(f"output/{rut}.docx")

        password = rut.split("-")[0].replace(".", "")

        # Encryptamos el pdf
        encrypt_pdf(f"output/{rut}.pdf", password)



if __name__ == "__main__":
    main()
