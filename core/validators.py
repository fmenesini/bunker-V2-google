# core/validators.py
import pandas as pd
import streamlit as st

class DataSanitizer:
    @staticmethod
    def sanitize_excel_import(df: pd.DataFrame, expected_columns: list = None) -> pd.DataFrame:
        """
        Camera di decontaminazione per i file Excel caricati dai commerciali.
        Se il file contiene dati tossici, solleva un'eccezione bloccante.
        """
        # Creiamo una copia per non sporcare il dataframe originale in memoria
        df_clean = df.copy()

        # 1. Check di base: Ci sono le colonne che ci aspettiamo?
        if expected_columns:
            missing = [col for col in expected_columns if col not in df_clean.columns]
            if missing:
                raise ValueError(
                    f"FORMATO ERRATO: Mancano le colonne {', '.join(missing)}. "
                    "Hai usato il Template ufficiale?"
                )

        # 2. La trappola per l'EAN (Il vero Cavallo di Troia)
        if 'EAN' in df_clean.columns:
            # Elimina eventuali ".0" se pandas ha letto l'EAN come float, e pialla gli spazi
            df_clean['EAN'] = df_clean['EAN'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
            # Il cecchino contro la notazione scientifica: se trova la "E" e il "+", spara.
            if df_clean['EAN'].str.contains(r'[eE]\+', regex=True).any():
                raise ValueError(
                    "ALLARME EAN CORROTTO: Excel ha trasformato i codici a barre in notazione "
                    "scientifica (es. 8.00E+12). Apri l'Excel, formatta l'intera colonna EAN come 'Testo', "
                    "salva e ripeti l'upload."
                )
                
            # Verifica che siano rimasti *solo* numeri (esclude lettere o simboli strani)
            invalid_eans = df_clean[~df_clean['EAN'].str.match(r'^\d+$') & (df_clean['EAN'] != 'nan')]['EAN'].tolist()
            if invalid_eans:
                raise ValueError(f"ALLARME EAN SPORCO: Trovati caratteri non numerici -> {invalid_eans[:3]}")

        # 3. Piallatura intensiva (Spazi e NaN)
        for col in df_clean.columns:
            # Se la colonna contiene testo (object)
            if df_clean[col].dtype == 'object':
                df_clean[col] = df_clean[col].astype(str).str.strip()
                # Rimettiamo a None ciò che Pandas ha trasformato nella stringa letterale 'nan'
                df_clean[col] = df_clean[col].replace('nan', None)
            
            # Se la colonna contiene numeri (float/int)
            elif pd.api.types.is_numeric_dtype(df_clean[col]):
                # Trasforma le celle lasciate vuote in 0.0, prevenendo crash nel database
                df_clean[col] = df_clean[col].fillna(0.0)

        return df_clean
