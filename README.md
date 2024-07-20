# Excel to Secure PDF

Este proyecto en Python convierte un archivo Excel a un documento PDF encriptado utilizando un template de DOCX.

## Requisitos

- Python 3.x
- [Virtualenv](https://virtualenv.pypa.io/en/latest/)

## Instalación

1. **Clona el repositorio:**

    ```bash
    git clone <URL_DEL_REPOSITORIO>
    cd <NOMBRE_DEL_REPOSITORIO>
    ```

2. **Crea un entorno virtual:**

    ```bash
    virtualenv env
    ```
    ó
    ```bash
    python -m venv env
    ```

3. **Activa el entorno virtual:**

    - En Windows:

      ```bash
      env\Scripts\activate
      ```

    - En macOS/Linux:

      ```bash
      source env/bin/activate
      ```

4. **Instala las dependencias:**

    ```bash
    pip install -r requirements.txt
    ```

## Uso

1. **Prepara tu archivo Excel y el template DOCX.**

2. **Ejecuta el script:**

    ```bash
    python src/main.py
    ```

## Descripción del Script

El script realiza las siguientes acciones:

1. **Lee el archivo Excel**: Se extraen los datos necesarios.

2. **Reemplaza valores en el template DOCX**: Los datos del Excel se insertan en el template DOCX.

3. **Convierte el documento DOCX a PDF**: Utiliza herramientas de conversión para obtener un PDF del documento.

4. **Encripta el PDF**: El PDF resultante se encripta con una contraseña especificada.

## Contribuciones

Las contribuciones son bienvenidas. Por favor, abre un _issue_ o un _pull request_ si deseas mejorar el proyecto.

## Licencia

Este proyecto está licenciado bajo la [Licencia MIT](LICENSE).
