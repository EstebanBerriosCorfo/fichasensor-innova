from src.utils.TextUtils import TextUtils
import json

class ProjectInfo:
    """
    Modelo que representa un proyecto y su historial de reuniones/visitas técnicas.
    Aplica formato Título a 'Representante Legal' (SOAP) y a 'Nombre' y 'Responsable' (Excel).
    """

    def __init__(self, soap_data, meetings):
        # Clona el dict de soap_data y aplica título a 'Representante Legal'
        self.project = dict(soap_data) if soap_data else {}
        if "Representante Legal" in self.project:
            self.project["Representante Legal"] = TextUtils.format_title_case(self.project["Representante Legal"])

        # Prepara la lista de reuniones con formateo en campos relevantes
        self.meetings = []
        for m in meetings or []:
            fila = dict(m)  # copia defensiva
            # Formato título para 'Nombre' y 'Responsable'
            if "Nombre" in fila:
                fila["Nombre"] = TextUtils.format_title_case(fila["Nombre"])
            if "Responsable" in fila:
                fila["Responsable"] = TextUtils.format_title_case(fila["Responsable"])
            # Correos en minúscula usando TextUtils
            if "Correo electrónico" in fila:
                fila["Correo electrónico"] = TextUtils.to_lower(fila["Correo electrónico"])
            if "Correo electrónico de contacto" in fila:
                fila["Correo electrónico de contacto"] = TextUtils.to_lower(fila["Correo electrónico de contacto"])
            self.meetings.append(fila)

    def to_dict(self):
        """
        Retorna el dict estructurado, listo para serializar.
        """
        return {
            "project": self.project,
            "meetings": self.meetings
        }

    def to_json(self, **kwargs):
        """
        Serializa el objeto a JSON legible.
        :param kwargs: argumentos opcionales para json.dumps (indent, ensure_ascii, etc.)
        """
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2, **kwargs)
