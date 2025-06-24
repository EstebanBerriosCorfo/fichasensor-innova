import pandas as pd
from src.utils.FileUtils import FileUtils

class ExcelReader:
    def __init__(self):
        excel_path = FileUtils.get_fichasensor_excel_path()
        # Especifica explícitamente la hoja 'Sheet1'
        self.dataFrame = pd.read_excel(excel_path, sheet_name="Sheet1")

    def get_project_rows(self, projectCode):
        df = self.dataFrame
        # Filtro robusto, trae todas las filas que coincidan
        result = df[df['Código proyecto'].astype(str).str.strip().str.upper() == projectCode.strip().upper()]
        if not result.empty:
            return result.to_dict(orient='records')
        return []
    
    def get_all_codes(self):
        """
        Obtiene todos los códigos de proyecto únicos del DataFrame.
        :return: Lista de códigos de proyecto únicos
        """
        return self.dataFrame['Código proyecto'].dropna().unique().tolist()

if __name__ == "__main__":
    project_code = input("Ingrese el código de proyecto a buscar: ")
    try:
        reader = ExcelReader()
        project_data = reader.get_project_row(project_code)
        if project_data:
            print("Datos encontrados para el proyecto:")
            for k, v in project_data.items():
                print(f"{k}: {v}")
        else:
            print("No se encontró el proyecto.")
            print("Columnas disponibles en el archivo:", reader.dataFrame.columns.tolist())
    except FileNotFoundError as e:
        print(str(e))
