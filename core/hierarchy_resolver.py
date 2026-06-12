# core/hierarchy_resolver.py
import sqlite3
from decimal import Decimal
from typing import Optional
from dataclasses import dataclass

@dataclass
class ResolvedContract:
    listino_r: Optional[Decimal] = None
    sconto_1: Optional[Decimal] = None
    sconto_2: Optional[Decimal] = None
    sconto_3: Optional[Decimal] = None
    sconto_4: Optional[Decimal] = None
    sconto_5: Optional[Decimal] = None
    sconto_6: Optional[Decimal] = None
    sconto_7: Optional[Decimal] = None
    sconto_y: Optional[Decimal] = None
    sconto_carico: Optional[Decimal] = None
    sconto_pagamento: Optional[Decimal] = None
    voce_i: Optional[Decimal] = None
    voce_ii: Optional[Decimal] = None
    voce_iii: Optional[Decimal] = None
    voce_iv: Optional[Decimal] = None
    voce_v: Optional[Decimal] = None
    livello_risolto: str = "NESSUNO"

class HierarchyResolver:
    _FIELDS = {
        "sconto_1": "sconto_1", "sconto_2": "sconto_2", "sconto_3": "sconto_3",
        "sconto_4": "sconto_4", "sconto_5": "sconto_5", "sconto_6": "sconto_6",
        "sconto_7": "sconto_7", "sconto_y": "sconto_y", "sconto_carico": "sconto_carico", 
        "sconto_pagamento": "sconto_pagamento", "voce_contratto_1": "voce_i",
        "voce_contratto_2": "voce_ii", "voce_contratto_3": "voce_iii",
        "voce_contratto_4": "voce_iv", "voce_contratto_5": "voce_v"
    }

    @classmethod
    def resolve(cls, conn: sqlite3.Connection, gruppo: str, sottogruppo: str, insegna: str, ean: str, categoria: str) -> ResolvedContract:
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # Ordinamento Bottom-Up: elabora prima i livelli più specifici (REFERENZA -> INSEGNA -> CATEGORIA -> SOTTOGRUPPO -> GRUPPO)
        cursor.execute("""
            SELECT livello, chiave_livello, listino_r,
                   sconto_1, sconto_2, sconto_3, sconto_4, sconto_5,
                   sconto_6, sconto_7, sconto_y, sconto_carico, sconto_pagamento,
                   voce_contratto_1, voce_contratto_2, voce_contratto_3, 
                   voce_contratto_4, voce_contratto_5
            FROM accordi_commerciali
            WHERE (UPPER(TRIM(gruppo_macro)) = UPPER(TRIM(?)) AND livello = 'GRUPPO' AND (sottogruppo = '' OR sottogruppo IS NULL) AND (associato_insegna = '' OR associato_insegna IS NULL))
               OR (UPPER(TRIM(gruppo_macro)) = UPPER(TRIM(?)) AND UPPER(TRIM(sottogruppo)) = UPPER(TRIM(?)) AND livello = 'SOTTOGRUPPO' AND (associato_insegna = '' OR associato_insegna IS NULL))
               OR (UPPER(TRIM(gruppo_macro)) = UPPER(TRIM(?)) AND UPPER(TRIM(sottogruppo)) = UPPER(TRIM(?)) AND UPPER(TRIM(associato_insegna)) = UPPER(TRIM(?)) AND livello = 'CATEGORIA' AND UPPER(TRIM(chiave_livello)) = UPPER(TRIM(?)))
               OR (UPPER(TRIM(gruppo_macro)) = UPPER(TRIM(?)) AND UPPER(TRIM(sottogruppo)) = UPPER(TRIM(?)) AND UPPER(TRIM(associato_insegna)) = UPPER(TRIM(?)) AND livello = 'INSEGNA')
               OR (UPPER(TRIM(gruppo_macro)) = UPPER(TRIM(?)) AND UPPER(TRIM(sottogruppo)) = UPPER(TRIM(?)) AND UPPER(TRIM(associato_insegna)) = UPPER(TRIM(?)) AND livello = 'REFERENZA' AND chiave_livello = ?)
            ORDER BY 
                CASE livello
                    WHEN 'REFERENZA' THEN 1
                    WHEN 'INSEGNA' THEN 2
                    WHEN 'CATEGORIA' THEN 3
                    WHEN 'SOTTOGRUPPO' THEN 4
                    WHEN 'GRUPPO' THEN 5
                END ASC
        """, (gruppo, 
              gruppo, sottogruppo, 
              gruppo, sottogruppo, insegna, categoria, 
              gruppo, sottogruppo, insegna,
              gruppo, sottogruppo, insegna, ean))

        rules = cursor.fetchall()
        
        locked_values = {attr: None for attr in cls._FIELDS.values()}
        locked_values['listino_r'] = None
        livello_risolto = "NESSUNO"

        for row in rules:
            liv = row["livello"]
            chiave = (row["chiave_livello"] or "").upper().strip()

            if liv == "CATEGORIA" and chiave != categoria.upper().strip():
                continue
            if liv == "REFERENZA" and chiave != ean.strip():
                continue

            livello_risolto = liv

            # Il primo record specifico che valorizza il parametro ne blocca la riscrittura da parte di record generali
            if locked_values['listino_r'] is None and row["listino_r"] is not None:
                locked_values['listino_r'] = Decimal(str(row["listino_r"]))

            for db_field, attr in cls._FIELDS.items():
                val = row[db_field]
                if val is not None and locked_values[attr] is None:
                    locked_values[attr] = Decimal(str(val))

        return ResolvedContract(
            listino_r=locked_values['listino_r'],
            sconto_1=locked_values['sconto_1'],
            sconto_2=locked_values['sconto_2'],
            sconto_3=locked_values['sconto_3'],
            sconto_4=locked_values['sconto_4'],
            sconto_5=locked_values['sconto_5'],
            sconto_6=locked_values['sconto_6'],
            sconto_7=locked_values['sconto_7'],
            sconto_y=locked_values['sconto_y'],
            sconto_carico=locked_values['sconto_carico'],
            sconto_pagamento=locked_values['sconto_pagamento'],
            voce_i=locked_values['voce_i'],
            voce_ii=locked_values['voce_ii'],
            voce_iii=locked_values['voce_iii'],
            voce_iv=locked_values['voce_iv'],
            voce_v=locked_values['voce_v'],
            livello_risolto=livello_risolto
        )
