import streamlit as st
import sqlite3
import pandas as pd
import random
from decimal import Decimal
from datetime import date
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, GridUpdateMode, DataReturnMode, ColumnsAutoSizeMode
from config import DB_FILE
from core.hierarchy_resolver import HierarchyResolver

class RenewalEngine:
    @staticmethod
    def calculate_rollups(df_input):
        """
        Motore di calcolo backend per le medie ponderate.
        Calcola i valori per Referenza, Sub-Categoria, Categoria e Totale.
        """
        df = df_input.copy()
        
        # 1. Calcoli Base per Referenza (SKU)
        df['Net_Net_N'] = df['[N] Listino €'] * (1 - (df['[N] Sc. Fattura %'] / 100)) * (1 - (df['[N] PFA %'] / 100))
        df['Fatturato_N'] = df['Net_Net_N'] * df['[N] Volumi']
        
        df['Net_Net_N1'] = df['[N+1] Listino €'] * (1 - (df['[N+1] Sc. Fattura %'] / 100)) * (1 - (df['[N+1] PFA %'] / 100))
        df['Fatturato_N1'] = df['Net_Net_N1'] * df['[N+1] Volumi']
        
        df['Valore_Floor_N1'] = df['Floor Minimo €'] * df['[N+1] Volumi']
        
        # Delta % e Spazio Promo per SKU
        df['Delta_%'] = df.apply(lambda x: ((x['Net_Net_N1'] - x['Net_Net_N']) / x['Net_Net_N'] * 100) if x['Net_Net_N'] > 0 else 0, axis=1)
        df['Spazio_Promo_€'] = df['Net_Net_N1'] - df['Floor Minimo €']
        df['Allarme'] = df['Net_Net_N1'] < df['Floor Minimo €']

        # 2. Roll-up per Sub-Categoria
        df_subcat = df.groupby('Sub-Categoria').agg(
            Volumi_N=('[N] Volumi', 'sum'),
            Fatturato_N=('Fatturato_N', 'sum'),
            Volumi_N1=('[N+1] Volumi', 'sum'),
            Fatturato_N1=('Fatturato_N1', 'sum'),
            Valore_Floor_N1=('Valore_Floor_N1', 'sum')
        ).reset_index()
        
        df_subcat['Net_Net_Pond_N'] = df_subcat.apply(lambda x: x['Fatturato_N'] / x['Volumi_N'] if x['Volumi_N'] > 0 else 0, axis=1)
        df_subcat['Net_Net_Pond_N1'] = df_subcat.apply(lambda x: x['Fatturato_N1'] / x['Volumi_N1'] if x['Volumi_N1'] > 0 else 0, axis=1)
        df_subcat['Floor_Pond_N1'] = df_subcat.apply(lambda x: x['Valore_Floor_N1'] / x['Volumi_N1'] if x['Volumi_N1'] > 0 else 0, axis=1)
        df_subcat['Delta_%'] = df_subcat.apply(lambda x: ((x['Net_Net_Pond_N1'] - x['Net_Net_Pond_N']) / x['Net_Net_Pond_N'] * 100) if x['Net_Net_Pond_N'] > 0 else 0, axis=1)
        df_subcat['Allarme'] = df_subcat['Net_Net_Pond_N1'] < df_subcat['Floor_Pond_N1']

        # 3. Roll-up per Categoria Macro
        df_cat = df.groupby('Categoria').agg(
            Volumi_N=('[N] Volumi', 'sum'),
            Fatturato_N=('Fatturato_N', 'sum'),
            Volumi_N1=('[N+1] Volumi', 'sum'),
            Fatturato_N1=('Fatturato_N1', 'sum'),
            Valore_Floor_N1=('Valore_Floor_N1', 'sum')
        ).reset_index()
        
        df_cat['Net_Net_Pond_N'] = df_cat.apply(lambda x: x['Fatturato_N'] / x['Volumi_N'] if x['Volumi_N'] > 0 else 0, axis=1)
        df_cat['Net_Net_Pond_N1'] = df_cat.apply(lambda x: x['Fatturato_N1'] / x['Volumi_N1'] if x['Volumi_N1'] > 0 else 0, axis=1)
        df_cat['Floor_Pond_N1'] = df_cat.apply(lambda x: x['Valore_Floor_N1'] / x['Volumi_N1'] if x['Volumi_N1'] > 0 else 0, axis=1)
        df_cat['Delta_%'] = df_cat.apply(lambda x: ((x['Net_Net_Pond_N1'] - x['Net_Net_Pond_N']) / x['Net_Net_Pond_N'] * 100) if x['Net_Net_Pond_N'] > 0 else 0, axis=1)
        df_cat['Allarme'] = df_cat['Net_Net_Pond_N1'] < df_cat['Floor_Pond_N1']

        # 4. Totali Globali
        tot_vol_n = df['[N] Volumi'].sum()
        tot_vol_n1 = df['[N+1] Volumi'].sum()
        tot_fatt_n = df['Fatturato_N'].sum()
        tot_fatt_n1 = df['Fatturato_N1'].sum()
        tot_floor_n1 = df['Valore_Floor_N1'].sum()

        totali = {
            'Volumi_N1': tot_vol_n1,
            'Net_Net_Pond_N': tot_fatt_n / tot_vol_n if tot_vol_n > 0 else 0,
            'Net_Net_Pond_N1': tot_fatt_n1 / tot_vol_n1 if tot_vol_n1 > 0 else 0,
            'Floor_Pond_N1': tot_floor_n1 / tot_vol_n1 if tot_vol_n1 > 0 else 0
        }
        totali['Delta_%'] = ((totali['Net_Net_Pond_N1'] - totali['Net_Net_Pond_N']) / totali['Net_Net_Pond_N'] * 100) if totali['Net_Net_Pond_N'] > 0 else 0

        return df, df_subcat, df_cat, totali

def render_simulazione_rinnovi():
    st.title("Simulazione Rinnovi Contrattuali (N vs N+1)")
    st.markdown("Analisi differenziale dei margini, calcolo dello Spazio Promo e Roll-up per Categorie e Sub-Categorie.")
    
    anno_corrente = date.today().year
    conn = sqlite3.connect(DB_FILE)
    
    # --- 1. CONTESTO E PRE-COMPILAZIONE ---
    st.markdown("#### 1. Contesto di Riferimento (Pre-compilazione Anno N)")
    
    with st.container(border=True):
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT gruppo_macro FROM clienti WHERE attivo=1 ORDER BY gruppo_macro")
        gruppi = [r[0] for r in cursor.fetchall()]
        
        col_ctx1, col_ctx2 = st.columns(2)
        with col_ctx1:
            gruppo_sel = st.selectbox("Gruppo GDO", ["Nessuno"] + gruppi)
        
        sottogruppi = []
        if gruppo_sel != "Nessuno":
            cursor.execute("SELECT DISTINCT sottogruppo FROM clienti WHERE gruppo_macro=? AND attivo=1 ORDER BY sottogruppo", (gruppo_sel,))
            sottogruppi = [r[0] for r in cursor.fetchall()]
        with col_ctx2:
            sottogruppo_sel = st.selectbox("Sottogruppo GDO", [""] + sottogruppi if gruppo_sel != "Nessuno" else [""])
            
        associato_sel = "" # Fermo al sottogruppo per i rinnovi

    # Estrazione Anagrafica e Classificazione
    query = """
        SELECT a.ean, a.descrizione_commerciale, a.tipo_olio, COALESCE(g.min_net_net_g, 0.0) as min_net_net_g
        FROM anagrafica_master a
        LEFT JOIN guardrail_aziendali g ON a.ean = g.ean
    """
    df_base = pd.read_sql_query(query, conn)
    
    def get_subcat(row):
        desc = str(row['descrizione_commerciale']).upper()
        tipo = str(row['tipo_olio']).upper()
        if tipo == 'EXTRAVERGINE':
            if '100% ITA' in desc or '100%I' in desc or 'TOSC' in desc: return 'Extravergini Italiani'
            if 'BIO' in desc: return 'Extravergini Biologici'
            return 'Extravergini Comunitari'
        elif tipo == 'OLIVA': return 'Oli di Oliva Raffinati'
        elif tipo == 'SEMI':
            if 'ARACHIDE' in desc: return 'Semi di Arachide'
            if 'MAIS' in desc: return 'Semi di Mais'
            if 'GIRAS' in desc: return 'Semi di Girasole'
            if 'FRITT' in desc or 'FRIMX' in desc: return 'Oli per Frittura Specifici'
            if 'VINACC' in desc: return 'Semi di Vinacciolo'
            return 'Altri Oli di Semi'
        elif tipo == 'ACETO': return 'Aceto Balsamico'
        return 'Altro'
        
    df_base['Sub-Categoria'] = df_base.apply(get_subcat, axis=1)
    df_base['Categoria'] = df_base['tipo_olio']
    df_base = df_base.rename(columns={'descrizione_commerciale': 'Prodotto', 'min_net_net_g': 'Floor Minimo €'})
    
    operative_cols = [
        '[N] Volumi', '[N] Listino €', '[N] Sc. Fattura %', '[N] PFA %',
        '[N+1] Volumi', '[N+1] Listino €', '[N+1] Sc. Fattura %', '[N+1] PFA %'
    ]
    for col in operative_cols:
        df_base[col] = 0.0 if '€' in col or '%' in col else 0

    # Inizializzazione Session State
    if 'rinnovi_df' not in st.session_state:
        st.session_state.rinnovi_df = df_base.copy()

    # Pulsanti di Azione
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("Carica Condizioni Attuali da DB", type="primary", use_container_width=True):
            if gruppo_sel != "Nessuno":
                df_temp = st.session_state.rinnovi_df.copy()
                for idx, row in df_temp.iterrows():
                    contract = HierarchyResolver.resolve(conn, gruppo_sel, sottogruppo_sel, associato_sel, row['ean'], row['tipo_olio'])
                    if contract.listino_r is not None:
                        df_temp.at[idx, '[N] Listino €'] = float(contract.listino_r)
                        p = contract.listino_r
                        for s in [contract.sconto_1, contract.sconto_2, contract.sconto_3, contract.sconto_4, contract.sconto_5, contract.sconto_6, contract.sconto_7, contract.sconto_y, contract.sconto_carico, contract.sconto_pagamento]:
                            p = p * (Decimal('1') - (s / Decimal('100')))
                        if contract.listino_r > 0:
                            sc_fatt_eq = (Decimal('1') - (p / contract.listino_r)) * Decimal('100')
                            df_temp.at[idx, '[N] Sc. Fattura %'] = float(sc_fatt_eq.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
                        pfa_tot = contract.voce_i + contract.voce_ii + contract.voce_iii + contract.voce_iv + contract.voce_v
                        df_temp.at[idx, '[N] PFA %'] = float(pfa_tot)
                        # Copia i volumi N in N+1 di default (inizializzati a 0, ma pronti per l'input)
                st.session_state.rinnovi_df = df_temp
                st.rerun()
            else:
                st.warning("Seleziona un Gruppo GDO prima di caricare i dati.")

    with col_btn2:
        if st.button("Popola con Dati di Test (Mock Data)", use_container_width=True):
            df_mock = st.session_state.rinnovi_df.copy()
            for idx, row in df_mock.iterrows():
                floor = row['Floor Minimo €'] if row['Floor Minimo €'] > 0 else 3.0
                vol = random.randint(10, 100) * 100
                df_mock.at[idx, '[N] Volumi'] = vol
                df_mock.at[idx, '[N] Listino €'] = float(round(floor * 1.5, 2))
                df_mock.at[idx, '[N] Sc. Fattura %'] = 10.0
                df_mock.at[idx, '[N] PFA %'] = 5.0
                df_mock.at[idx, '[N+1] Volumi'] = int(vol * 1.05) # Incremento volumi 5%
                df_mock.at[idx, '[N+1] Listino €'] = float(round(floor * 1.6, 2)) # Aumento listino
                df_mock.at[idx, '[N+1] Sc. Fattura %'] = 12.0
                df_mock.at[idx, '[N+1] PFA %'] = 5.0
            st.session_state.rinnovi_df = df_mock
            st.rerun()

    conn.close()

    # --- TABS ---
    tab_simulazione, tab_risultati = st.tabs([
        "1. Master Grid (Input Dati)", 
        "2. Analisi Ponderata & Spazio Promo"
    ])

    with tab_simulazione:
        st.markdown("#### Griglia di Simulazione Contrattuale")
        st.markdown("Modifica i dati direttamente nella tabella. I calcoli nel Tab 'Analisi' si aggiorneranno automaticamente.")
        
        df_for_grid = st.session_state.rinnovi_df[['ean', 'Categoria', 'Sub-Categoria', 'Prodotto', 'Floor Minimo €'] + operative_cols]
        
        gb = GridOptionsBuilder.from_dataframe(df_for_grid)
        
        # Raggruppamento e blocco colonne
        gb.configure_column("Categoria", rowGroup=True, hide=True)
        gb.configure_column("Sub-Categoria", rowGroup=True, hide=True)
        gb.configure_column("Prodotto", pinned='left', width=250)
        gb.configure_column("ean", hide=True)
        gb.configure_column("Floor Minimo €", type=["numericColumn"], valueFormatter="x.toFixed(2) + ' €'", width=120)
        
        # Formattazione e modifica colonne operative
        for col in operative_cols:
            if '€' in col:
                gb.configure_column(col, editable=True, type=["numericColumn"], valueFormatter="x.toFixed(2) + ' €'", width=110)
            elif '%' in col:
                gb.configure_column(col, editable=True, type=["numericColumn"], valueFormatter="x.toFixed(2) + ' %'", width=110)
            else:
                gb.configure_column(col, editable=True, type=["numericColumn"], width=100)

        # Iniezione JS per calcolo Live del Net-Net N+1 nella griglia
        net_net_jscode = JsCode('''
        function(params) {
            let data = params.data;
            if (!data) return 0;
            let listino = data['[N+1] Listino €'] || 0;
            let sc = data['[N+1] Sc. Fattura %'] || 0;
            let pfa = data['[N+1] PFA %'] || 0;
            let netto_fattura = listino * (1 - (sc / 100));
            return netto_fattura * (1 - (pfa / 100));
        }
        ''')
        gb.configure_column("Net_Net_Live", header_name="Net-Net [N+1] €", valueGetter=net_net_jscode, type=["numericColumn"], valueFormatter="x.toFixed(3) + ' €'", width=130)

        gb.configure_grid_options(
            enableRangeSelection=True,
            suppressAggFuncInHeader=True,
            groupDefaultExpanded=-1,
            domLayout='normal'
        )

        gridOptions = gb.build()

        grid_response = AgGrid(
            df_for_grid,
            gridOptions=gridOptions,
            update_mode=GridUpdateMode.VALUE_CHANGED, # Aggiorna il backend ad ogni modifica
            data_return_mode=DataReturnMode.AS_INPUT,
            allow_unsafe_jscode=True,
            theme='alpine',
            height=600,
            fit_columns_on_grid_load=False
        )
        
        # Salvataggio modifiche nel Session State
        if grid_response['data'] is not None:
            df_returned = pd.DataFrame(grid_response['data'])
            for col in operative_cols:
                st.session_state.rinnovi_df[col] = df_returned[col]

    with tab_risultati:
        # Filtriamo solo le referenze con volumi N+1 > 0 per l'analisi
        df_active = st.session_state.rinnovi_df[st.session_state.rinnovi_df['[N+1] Volumi'] > 0].copy()
        
        if df_active.empty:
            st.warning("Nessuna referenza attiva. Inserisci dei volumi nella colonna '[N+1] Volumi' nella Master Grid.")
        else:
            # Esecuzione Motore Python per i Roll-up
            df_sku, df_subcat, df_cat, totali = RenewalEngine.calculate_rollups(df_active)
            
            # --- KPI GLOBALI ---
            st.markdown("#### KPI Totali Cliente (Media Ponderata Globale)")
            col_k1, col_k2, col_k3, col_k4 = st.columns(4)
            col_k1.metric("Volumi Totali [N+1]", f"{totali['Volumi_N1']:,.0f} Pz")
            col_k2.metric("Net-Net Pond. [N]", f"€ {totali['Net_Net_Pond_N']:.3f}")
            col_k3.metric("Net-Net Pond. [N+1]", f"€ {totali['Net_Net_Pond_N1']:.3f}", f"{totali['Net_Net_Pond_N1'] - totali['Net_Net_Pond_N']:+.3f} € vs [N]")
            col_k4.metric("Variazione Totale (%)", f"{totali['Delta_%']:+.2f} %", "Delta % vs Anno N")
            
            st.divider()
            
            # Funzione di stile per gli allarmi
            def highlight_alarms(row):
                if row['Allarme']: return ['background-color: #FEF2F2; color: #991B1B; font-weight: bold'] * len(row)
                if row['Delta_%'] < 0: return ['color: #D97706'] * len(row) # Arancione se perdo % ma sono sopra floor
                return [''] * len(row)

            # --- ROLL-UP CATEGORIA ---
            st.markdown("#### 1. Sintesi per Categoria Macro")
            df_cat_disp = df_cat[['Categoria', 'Volumi_N1', 'Net_Net_Pond_N', 'Net_Net_Pond_N1', 'Delta_%', 'Floor_Pond_N1', 'Allarme']]
            st.dataframe(
                df_cat_disp.style.apply(highlight_alarms, axis=1).format({
                    'Net_Net_Pond_N': '€ {:.3f}', 'Net_Net_Pond_N1': '€ {:.3f}', 
                    'Delta_%': '{:+.2f} %', 'Floor_Pond_N1': '€ {:.3f}'
                }),
                column_config={"Allarme": "Sotto Floor!"},
                hide_index=True, use_container_width=True
            )
            
            st.markdown("<br>", unsafe_allow_html=True)

            # --- ROLL-UP SUB-CATEGORIA ---
            st.markdown("#### 2. Sintesi per Sub-Categoria")
            df_subcat_disp = df_subcat[['Sub-Categoria', 'Volumi_N1', 'Net_Net_Pond_N', 'Net_Net_Pond_N1', 'Delta_%', 'Floor_Pond_N1', 'Allarme']]
            st.dataframe(
                df_subcat_disp.style.apply(highlight_alarms, axis=1).format({
                    'Net_Net_Pond_N': '€ {:.3f}', 'Net_Net_Pond_N1': '€ {:.3f}', 
                    'Delta_%': '{:+.2f} %', 'Floor_Pond_N1': '€ {:.3f}'
                }),
                column_config={"Allarme": "Sotto Floor!"},
                hide_index=True, use_container_width=True
            )
            
            st.divider()
            
            # --- DETTAGLIO REFERENZE (SKU) ---
            st.markdown("#### 3. Dettaglio Referenze (SKU) e Spazio Promo")
            st.markdown("La colonna **Spazio Promo (€)** indica il margine unitario residuo prima di andare in perdita sul Floor.")
            
            cols_sku_disp = ['Sub-Categoria', 'Prodotto', 'Net_Net_N', 'Net_Net_N1', 'Delta_%', 'Floor Minimo €', 'Spazio_Promo_€', 'Allarme']
            st.dataframe(
                df_sku[cols_sku_disp].style.apply(highlight_alarms, axis=1).format({
                    'Net_Net_N': '€ {:.3f}', 'Net_Net_N1': '€ {:.3f}', 
                    'Delta_%': '{:+.2f} %', 'Floor Minimo €': '€ {:.3f}', 'Spazio_Promo_€': '€ {:+.3f}'
                }),
                column_config={"Allarme": "Sotto Floor!"},
                hide_index=True, use_container_width=True
            )
