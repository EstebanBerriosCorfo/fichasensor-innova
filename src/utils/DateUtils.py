from datetime import datetime, timedelta

class DateUtils:
    @staticmethod
    def format_date(date):
        """
        Devuelve un string con formato dd-mm-aaaa si la entrada es un datetime, pandas.Timestamp,
        un serial de fecha de Excel (float/int), o un string reconocible.
        Si la entrada no es reconocida como fecha, devuelve str(date).
        """
        # Soporte para float/int (Excel serial)
        if isinstance(date, (float, int)):
            # Excel base date: 1899-12-30
            try:
                base_date = datetime(1899, 12, 30)
                real_date = base_date + timedelta(days=int(date))
                return real_date.strftime('%d-%m-%Y')
            except Exception:
                pass

        # Soporte para pandas.Timestamp
        try:
            import pandas as pd
            if isinstance(date, pd.Timestamp):
                return date.strftime('%d-%m-%Y')
        except ImportError:
            pass  # Si pandas no está instalado, ignora esta comprobación

        # Soporte para datetime.datetime
        if isinstance(date, datetime):
            return date.strftime('%d-%m-%Y')

        # Si es string y parece una fecha, intenta parsear
        if isinstance(date, str):
            for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
                try:
                    dt = datetime.strptime(date, fmt)
                    return dt.strftime('%d-%m-%Y')
                except ValueError:
                    continue  # Prueba el siguiente formato
            # Si falla, intenta ISO format
            try:
                dt = datetime.fromisoformat(date)
                return dt.strftime('%d-%m-%Y')
            except Exception:
                pass  # Si tampoco es ISO, sigue

        # Por defecto, devuelve string vacío si está vacío o nan
        if not date or str(date).lower() == 'nan':
            return ""
        return str(date)