from src.utils.DateUtils import DateUtils

class TextUtils:
    @staticmethod
    def format_title_case(text):
        """
        Devuelve el texto en formato Título (Title Case).
        Si el valor no es string, lo devuelve igual.
        """
        if isinstance(text, str):
            return text.title()
        return text
    
    @staticmethod
    def to_lower(text):
        """
        Devuelve el texto en minúsculas.
        Si el valor no es string, lo devuelve igual.
        """
        if isinstance(text, str):
            return text.lower()
        return text
    
    @staticmethod
    def to_safe_str(value):
        """
        Convierte cualquier valor a string seguro para Word:
        - Si es None, NaN, NaT, retorna string vacío.
        - Si es fecha, usa DateUtils.format_date().
        - Si es número, lo convierte a string.
        - Si es string, lo retorna tal cual.
        """
        import pandas as pd
        import numpy as np
        from datetime import datetime

        if value is None:
            return ""
        if isinstance(value, float) and (np.isnan(value) or str(value).lower() == "nan"):
            return ""
        # pandas.NaT
        if hasattr(pd, "_libs") and hasattr(pd._libs.tslibs, "nattype"):
            if isinstance(value, pd._libs.tslibs.nattype.NaTType):
                return ""
        # Fechas pandas/datetime
        if isinstance(value, (pd.Timestamp, datetime)):
            return DateUtils.format_date(value)
        return str(value)