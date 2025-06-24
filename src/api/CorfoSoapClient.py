from zeep import Client
from zeep.helpers import serialize_object

class CorfoSoapClient:
    def __init__(self, wsdl_url='http://osblb2.corfo.cl/OSB/PX000451_ConsultaSnapshotSGP?wsdl'):
        self.client = Client(wsdl=wsdl_url)

    def get_project_data(self, projectCode):
        request = {
            'GERENCIA': '',
            'INSTRUMENTO': '',
            'EVENTO': '',
            'PROYECTO': projectCode
        }
        try:
            response = self.client.service.SEL_SNAPSHOT_PROYECTOS(**request)
            # Procesa y retorna la respuesta ya parseada como dict
            response_serialized = serialize_object(response)
            return CorfoSoapClient.parse_soap_response(response_serialized)
        except Exception as e:
            print(f"Error SOAP: {e}")
            return None

    @staticmethod
    def parse_soap_response(response):
        datos = {}
        if not response:
            return datos
        record = response[0]
        rows = record.get('Row', [])
        if not rows:
            return datos
        columns = rows[0].get('Column', [])
        for col in columns:
            nombre = col.get('name')
            valor = col.get('_value_1')
            if nombre:
                datos[nombre] = valor
        return datos