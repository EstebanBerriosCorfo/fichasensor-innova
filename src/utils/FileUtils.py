import os

class FileUtils:
    @staticmethod
    def get_downloads_folder():
        """
        Función que obtiene la ruta de la carpeta de descargas del usuario actual.
        Si la carpeta de descargas no existe, se crea.
        """
        return os.path.join(os.path.expanduser('~'), 'Downloads')

    @staticmethod
    def get_fichasensor_excel_path():
        """
        Retorna la ruta completa del archivo FichaSensor.xlsx ubicado en la carpeta OneDrive del usuario actual.
        Si el archivo no existe, lanza una excepción FileNotFoundError.
        """
        user_folder = os.path.expanduser('~')
        excel_path = os.path.join(
            user_folder,
            'OneDrive - corfo.cl',
            'Documentos - SUBDIRECCIÓN DE MEJORA CONTINUA',
            'Ficha Sensor.xlsx'
        )
        if os.path.exists(excel_path):
            return excel_path
        else:
            raise FileNotFoundError(f"No se encontró el archivo: {excel_path}")
        
    @staticmethod
    def get_template_path(template_name):
        """ Retorna la ruta absoluta de un archivo de plantilla específico.
        :param template_name: Nombre del archivo de plantilla (ej. 'template.html').
        :return: Ruta absoluta del archivo de plantilla.
        """
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # .../src/
        return os.path.join(base_dir, "templates", template_name)

    @staticmethod
    def get_assets_folder():
        """
        Retorna la ruta absoluta de la carpeta assets (src/assets).
        """
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # .../src/
        return os.path.join(base_dir, "assets")

    @staticmethod
    def get_logo_corfo_path():
        """
        Retorna la ruta absoluta de logoCorfo.png.
        """
        return os.path.join(FileUtils.get_assets_folder(), "logoCorfo.png")

    @staticmethod
    def get_icon_corfo_path():
        """
        Retorna la ruta absoluta de Corfo.jpg (ícono app).
        """
        return os.path.join(FileUtils.get_assets_folder(), "Corfo.jpg")
    
    @staticmethod
    def get_ico_corfo_path():
        """
        Retorna la ruta absoluta de corfo.ico (ícono de la aplicación).
        """
        return os.path.join(FileUtils.get_assets_folder(), "corfo.ico")