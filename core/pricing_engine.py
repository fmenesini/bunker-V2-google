# core/pricing_engine.py
from decimal import Decimal, ROUND_HALF_UP
from typing import List
from dataclasses import dataclass

@dataclass(frozen=True)
class PricingInput:
    listino_r: Decimal
    sconto_1: Decimal = Decimal("0.00")
    sconto_2: Decimal = Decimal("0.00")
    sconto_3: Decimal = Decimal("0.00")
    sconto_4: Decimal = Decimal("0.00")
    sconto_5: Decimal = Decimal("0.00")
    sconto_6: Decimal = Decimal("0.00")
    sconto_7: Decimal = Decimal("0.00")
    sconto_y: Decimal = Decimal("0.00")   # Sconto Continuativo
    sconto_z: Decimal = Decimal("0.00")   # Sconto Promozionale
    sconto_aa: Decimal = Decimal("0.00")  # Sconto Unitario in fattura
    sconto_carico: Decimal = Decimal("0.00")     # Logistica (AB)
    sconto_pagamento: Decimal = Decimal("0.00")  # Finanziario (AC)
    voce_i: Decimal = Decimal("0.00")     # PFA 1
    voce_ii: Decimal = Decimal("0.00")    # PFA 2
    voce_iii: Decimal = Decimal("0.00")   # PFA 3
    voce_iv: Decimal = Decimal("0.00")    # PFA 4
    voce_v: Decimal = Decimal("0.00")     # PFA 5 (Locale)
    min_net_net_g: Decimal = Decimal("0.00")  # Guardrail

@dataclass(frozen=True)
class WaterfallStep:
    fase: str
    valore: Decimal
    descrizione: str

@dataclass
class PricingResult:
    steps: List[WaterfallStep]
    netto_in_fattura_2: Decimal
    contratto_tot_pfa: Decimal
    net_net_finale: Decimal
    delta_vs_min: Decimal
    guardrail_ok: bool
    sconto_max_av: Decimal

class PricingEngine:
    @staticmethod
    def _apply_pct(base: Decimal, pct: Decimal) -> Decimal:
        if pct < Decimal("0.00") or pct >= Decimal("100.00"):
            raise ValueError(f"Percentuale sconto fuori limite consentito: {pct}%")
        factor = (Decimal("100.00") - pct) / Decimal("100.00")
        # Simula il calcolo gestionale: precisione a 5 decimali per ogni step intermedio
        return (base * factor).quantize(Decimal("0.00001"), rounding=ROUND_HALF_UP)

    @classmethod
    def calculate(cls, inp: PricingInput) -> PricingResult:
        steps: List[WaterfallStep] = []
        prezzo = inp.listino_r
        
        # 1. Listino
        steps.append(WaterfallStep("Listino Base (R)", prezzo, "Prezzo di partenza contrattuale"))

        # 2. Sconti Fissi Centrali (Esplosi singolarmente)
        if inp.sconto_1 != Decimal("0.00"):
            prezzo = cls._apply_pct(prezzo, inp.sconto_1)
            steps.append(WaterfallStep("Sconto Centrale 1", prezzo, f"-{inp.sconto_1}%"))
        if inp.sconto_2 != Decimal("0.00"):
            prezzo = cls._apply_pct(prezzo, inp.sconto_2)
            steps.append(WaterfallStep("Sconto Centrale 2", prezzo, f"-{inp.sconto_2}%"))
        if inp.sconto_3 != Decimal("0.00"):
            prezzo = cls._apply_pct(prezzo, inp.sconto_3)
            steps.append(WaterfallStep("Sconto Centrale 3", prezzo, f"-{inp.sconto_3}%"))
        if inp.sconto_4 != Decimal("0.00"):
            prezzo = cls._apply_pct(prezzo, inp.sconto_4)
            steps.append(WaterfallStep("Sconto Centrale 4", prezzo, f"-{inp.sconto_4}%"))
        if inp.sconto_5 != Decimal("0.00"):
            prezzo = cls._apply_pct(prezzo, inp.sconto_5)
            steps.append(WaterfallStep("Sconto Centrale 5", prezzo, f"-{inp.sconto_5}%"))

        # 3. Sconti Territoriali Locali (Esplosi singolarmente)
        if inp.sconto_6 != Decimal("0.00"):
            prezzo = cls._apply_pct(prezzo, inp.sconto_6)
            steps.append(WaterfallStep("Sconto Locale 6", prezzo, f"-{inp.sconto_6}%"))
        if inp.sconto_7 != Decimal("0.00"):
            prezzo = cls._apply_pct(prezzo, inp.sconto_7)
            steps.append(WaterfallStep("Sconto Locale 7", prezzo, f"-{inp.sconto_7}%"))

        # 4. Leve Venditore (Y & Z)
        if inp.sconto_y != Decimal("0.00"):
            prezzo = cls._apply_pct(prezzo, inp.sconto_y)
            steps.append(WaterfallStep("Sconto Continuativo (Y)", prezzo, f"-{inp.sconto_y}%"))

        if inp.sconto_z != Decimal("0.00"):
            prezzo = cls._apply_pct(prezzo, inp.sconto_z)
            steps.append(WaterfallStep("Sconto Promozionale (Z)", prezzo, f"-{inp.sconto_z}%"))

        # 5. Sconto Unitario in fattura (AA)
        if inp.sconto_aa != Decimal("0.00"):
            if inp.sconto_aa < Decimal("0.00"):
                raise ValueError("Lo sconto unitario non può essere negativo.")
            prezzo = max(Decimal("0.00"), prezzo - inp.sconto_aa)
            steps.append(WaterfallStep("Sconto Unitario in fattura (AA)", prezzo, f"-{inp.sconto_aa:.2f} Euro/Pz"))

        # 6. Oneri Logistici e Finanziari (AB & AC)
        if inp.sconto_carico != Decimal("0.00"):
            prezzo = cls._apply_pct(prezzo, inp.sconto_carico)
            steps.append(WaterfallStep("Oneri Logistica (AB)", prezzo, f"-{inp.sconto_carico}%"))
            
        if inp.sconto_pagamento != Decimal("0.00"):
            prezzo = cls._apply_pct(prezzo, inp.sconto_pagamento)
            steps.append(WaterfallStep("Oneri Pagamento (AC)", prezzo, f"-{inp.sconto_pagamento}%"))
        
        # Netto in Fattura 2
        netto_fatt_2 = prezzo.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)
        steps.append(WaterfallStep("Netto in Fattura 2 (AF)", netto_fatt_2, "Imponibile netto in fattura"))

        # 7. Premi Fuori Fattura (AL) - Somma algebrica con dettaglio voci
        pfa_sum = inp.voce_i + inp.voce_ii + inp.voce_iii + inp.voce_iv + inp.voce_v
        if pfa_sum >= Decimal("100.00"):
            raise ValueError(f"La somma dei Premi Fuori Fattura ({pfa_sum}%) supera o eguaglia il 100%. Ricavo impossibile.")
            
        if pfa_sum != Decimal("0.00"):
            if pfa_sum < Decimal("0.00"):
                raise ValueError("La somma dei PFA non può essere negativa.")
            pfa_details = []
            if inp.voce_i != Decimal("0.00"): pfa_details.append(f"I: {inp.voce_i}%")
            if inp.voce_ii != Decimal("0.00"): pfa_details.append(f"II: {inp.voce_ii}%")
            if inp.voce_iii != Decimal("0.00"): pfa_details.append(f"III: {inp.voce_iii}%")
            if inp.voce_iv != Decimal("0.00"): pfa_details.append(f"IV: {inp.voce_iv}%")
            if inp.voce_v != Decimal("0.00"): pfa_details.append(f"V: {inp.voce_v}%")
            
            net_net = cls._apply_pct(netto_fatt_2, pfa_sum)
            net_net = net_net.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)
            steps.append(WaterfallStep("Premi Fuori Fattura (AL)", net_net, f"Somma PFA: -{pfa_sum}% ({', '.join(pfa_details)})"))
        else:
            net_net = netto_fatt_2

        # 8. Net Net Finale
        steps.append(WaterfallStep("Net Net Finale (AM)", net_net, "Ricavo netto reale in cassa"))

        # Verifiche Guardrail con tolleranza al centesimo (-0.01) per evitare falsi positivi da rumore di arrotondamento
        delta = net_net - inp.min_net_net_g
        guardrail_ok = delta > Decimal("-0.01")

        # Sconto Massimo Scalare Teorico Totale (AV)
        sconto_max_av = Decimal("0.00")
        if inp.listino_r > Decimal("0.00"):
            pfa_factor = (Decimal("100.00") - pfa_sum) / Decimal("100.00")
            log_factor = (Decimal("100.00") - inp.sconto_carico) / Decimal("100.00")
            fin_factor = (Decimal("100.00") - inp.sconto_pagamento) / Decimal("100.00")
            combined = pfa_factor * log_factor * fin_factor
            if combined > Decimal("0.00"):
                prezzo_min_necessario = inp.min_net_net_g / combined
                prezzo_pre_aa_minimo = prezzo_min_necessario + inp.sconto_aa
                sconto_max_av_calc = (Decimal("1.00") - (prezzo_pre_aa_minimo / inp.listino_r)) * Decimal("100.00")
                sconto_max_av = max(Decimal("0.00"), sconto_max_av_calc).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

        return PricingResult(
            steps=steps,
            netto_in_fattura_2=netto_fatt_2,
            contratto_tot_pfa=pfa_sum,
            net_net_finale=net_net,
            delta_vs_min=delta,
            guardrail_ok=guardrail_ok,
            sconto_max_av=sconto_max_av
        )

    @classmethod
    def calculate_inverse(cls, target_net_net: Decimal, inp: PricingInput, target_field: str = "Z") -> Decimal:
        pfa_sum = inp.voce_i + inp.voce_ii + inp.voce_iii + inp.voce_iv + inp.voce_v
        
        if pfa_sum >= Decimal("100.00"):
            raise ValueError(f"Calcolo inverso impossibile: PFA totali al {pfa_sum}% distruggono il margine.")
            
        pfa_factor = (Decimal("100.00") - pfa_sum) / Decimal("100.00")
        netto_fatt_2_req = target_net_net / pfa_factor
        
        log_factor = (Decimal("100.00") - inp.sconto_carico) / Decimal("100.00")
        fin_factor = (Decimal("100.00") - inp.sconto_pagamento) / Decimal("100.00")
        log_fin_combined = log_factor * fin_factor
        
        if log_fin_combined <= Decimal("0.00"):
            raise ValueError("Oneri logistici/finanziari >= 100%. Impossibile calcolare.")
            
        prezzo_ante_log_req = netto_fatt_2_req / log_fin_combined
        prezzo_post_z_req = prezzo_ante_log_req + inp.sconto_aa

        prezzo_base = inp.listino_r
        for s in [inp.sconto_1, inp.sconto_2, inp.sconto_3, inp.sconto_4, inp.sconto_5, inp.sconto_6, inp.sconto_7, inp.sconto_y]:
            if s != Decimal("0.00"):
                prezzo_base = cls._apply_pct(prezzo_base, s)

        if target_field == "Z":
            if prezzo_base <= Decimal("0.00"):
                raise ValueError("Il prezzo base prima della Z è azzerato.")
            z_req = (Decimal("1.00") - (prezzo_post_z_req / prezzo_base)) * Decimal("100.00")
            return max(Decimal("0.00"), z_req).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

        elif target_field == "AA":
            prezzo_post_z = prezzo_base
            if inp.sconto_z != Decimal("0.00"):
                prezzo_post_z = cls._apply_pct(prezzo_post_z, inp.sconto_z)
            aa_req = prezzo_post_z - prezzo_ante_log_req
            return max(Decimal("0.00"), aa_req).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

        return Decimal("0.00")
