import json
from src.excel.ExcelReader import ExcelReader
from src.models.ProjectInfo import ProjectInfo

def main():
    codigo = input("Ingrese el código de proyecto: ").strip()
    reader = ExcelReader()
    filas = reader.get_project_rows(codigo)

    if not filas:
        print("No se encontraron filas para ese código de proyecto.")
        return

    # Simulación: soap_data (en tu flujo real deberías traerlo vía SOAP)
    soap_data = {
        "Nombre Proyecto": filas[0].get("Nombre Proyecto", ""),
        "Nombre Beneficiario": filas[0].get("Nombre Beneficiario", ""),
        "Representante Legal": filas[0].get("Representante Legal", ""),
        # ...otros campos relevantes
    }

    # Genera el objeto ProjectInfo (puedes personalizar los argumentos)
    project_info = ProjectInfo(filas[0], soap_data)
    
    # Si tu clase tiene to_json:
    if hasattr(project_info, "to_json"):
        print("\n📦 JSON serializado desde ProjectInfo:")
        print(project_info.to_json(indent=2, ensure_ascii=False))
    else:
        # Fallback si tienes to_dict:
        data = project_info.to_dict() if hasattr(project_info, "to_dict") else dict(project_info)
        print("\n📦 Dict serializado:")
        print(json.dumps(data, indent=2, ensure_ascii=False))

if __name__ == "__main__":
    main()
