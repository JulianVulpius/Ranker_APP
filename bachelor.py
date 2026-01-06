import streamlit as st
import pandas as pd
import numpy as np
from scipy.signal import welch
from scipy.stats import spearmanr, rankdata
import re
import altair as alt

# --------------------------
# KONSTANTEN & SETUP
# --------------------------
st.set_page_config(page_title="Bachelor EEG Analyse: Advanced", layout="wide")

SAMPLING_FREQUENCY = 250
EEG_CHANNELS = ["EEG 1", "EEG 2", "EEG 3", "EEG 4", "EEG 5", "EEG 6", "EEG 7", "EEG 8"]
CLEANING_LEVELS = ['Normal', 'Strikt', 'Sehr Strikt', 'Ohne']
METRICS = ['Alpha/Beta Verhältnis', 'Theta/Beta Verhältnis', '(Alpha+Theta)/Beta Verhältnis']

# --------------------------
# CONFIG: DEFAULT GENRES
# --------------------------
# Die Namen (Keys) müssen exakt mit den Songnamen in der CSV übereinstimmen.
DEFAULT_GENRES = {
    "Song 1": "Klassik",
    "Song 2": "Electronic",
    "Song 3": "Klassik",
    "Song 4": "Electronic",
    "Song 5": "Klassik",
    "Song 6": "Jazz",
    "Song 7": "EGitarre",
    "Song 8": "Gitarre",
    "Song 9": "Gitarre",
    "Song 10": "EGitarre"
}

# --------------------------
# SIGNAL VERARBEITUNG
# --------------------------

def compute_band_power(data_segment, sf):
    bands = {'Theta': (4, 8), 'Alpha': (8, 13), 'Beta': (13, 30)}
    missing_cols = [c for c in EEG_CHANNELS if c not in data_segment.columns]
    if missing_cols: return None
        
    eeg_data = data_segment[EEG_CHANNELS].values.T
    if eeg_data.shape[1] < sf: return None
    
    freqs, psd = welch(eeg_data, sf, nperseg=sf)
    res = {}
    for band, (low, high) in bands.items():
        idx = np.logical_and(freqs >= low, freqs <= high)
        res[band] = np.mean(psd[:, idx])
    return res

def calculate_ratios(band_power):
    if not band_power or band_power.get('Beta', 0) == 0:
        return {m: 0.0 for m in METRICS}
    a, t, b = band_power['Alpha'], band_power['Theta'], band_power['Beta']
    return {
        'Alpha/Beta Verhältnis': a / b,
        'Theta/Beta Verhältnis': t / b,
        '(Alpha+Theta)/Beta Verhältnis': (a + t) / b
    }

def get_global_events(df):
    if 'TRIG' not in df.columns: return [], [], []
    triggers = df[df['TRIG'] != 0]['TRIG'].astype(int).astype(str).tolist()
    indices = df[df['TRIG'] != 0].index.tolist()
    long_events, short_events, transitions = [], [], []
    pairs = {'3': '4', '5': '6', '7': '8', '9': '10'}
    processed = set()
    
    for i in range(len(triggers)):
        if indices[i] in processed: continue
        val = triggers[i]
        if val in pairs:
            target = pairs[val]
            for j in range(i+1, len(triggers)):
                if triggers[j] == target and indices[j] not in processed:
                    if val == '9': transitions.append((indices[i], indices[j]))
                    else: long_events.append((indices[i], indices[j]))
                    processed.add(indices[i]); processed.add(indices[j])
                    break
        elif val not in ['1', '2'] and val not in pairs.values():
            short_events.append(indices[i])
    return long_events, transitions, short_events

def clean_segment(segment, cleaning_level, short_radius_s, global_events):
    if cleaning_level == 'Ohne' or segment.empty: return segment
    longs, trans, shorts = global_events
    indices_to_drop = set()
    s_start, s_end = segment.index.min(), segment.index.max()
    def overlap(ev_start, ev_end): return max(ev_start, s_start) <= min(ev_end, s_end)

    if cleaning_level in ['Normal', 'Strikt', 'Sehr Strikt']:
        for t_s, t_e in trans:
            if overlap(t_s, t_e): indices_to_drop.update(segment.loc[max(t_s, s_start):min(t_e, s_end)].index)
    if cleaning_level in ['Strikt', 'Sehr Strikt']:
        for l_s, l_e in longs:
            if overlap(l_s, l_e): indices_to_drop.update(segment.loc[max(l_s, s_start):min(l_e, s_end)].index)
    if cleaning_level == 'Sehr Strikt':
        radius_idx = int(short_radius_s * SAMPLING_FREQUENCY)
        for sh_idx in shorts:
            if s_start <= sh_idx <= s_end:
                d_start = max(s_start, sh_idx - radius_idx)
                d_end = min(s_end, sh_idx + radius_idx)
                indices_to_drop.update(segment.loc[d_start:d_end].index)
    return segment.drop(index=list(indices_to_drop), errors='ignore')

def segment_song_data(df):
    if 'TRIG' not in df.columns: return []
    t = df[df['TRIG'].isin([1, 2])]
    if t.empty: return []
    segments = []
    starts = t[t['TRIG'] == 1].index
    master_end = df.index[-1]
    ends = t[t['TRIG'] == 2].index
    if not ends.empty: master_end = ends[-1]
    for start_idx in starts:
        if start_idx >= master_end: continue
        next_triggers = t[t.index > start_idx]
        end_idx = next_triggers.index[0] if not next_triggers.empty else master_end
        if end_idx > master_end: end_idx = master_end
        segments.append(df.loc[start_idx:end_idx])
    return segments

def load_eeg_file(file_obj, filename_context):
    file_obj.seek(0)
    try:
        df = pd.read_csv(file_obj)
        df.columns = df.columns.str.strip()
        if len(df.columns) < 2 and 'TRIG' not in df.columns:
            file_obj.seek(0); df = pd.read_csv(file_obj, sep=';'); df.columns = df.columns.str.strip()
        if 'TRIG' not in df.columns:
            found = False
            for col in df.columns:
                if 'trig' in col.lower(): df.rename(columns={col: 'TRIG'}, inplace=True); found = True; break
            if not found: raise ValueError("Spalte 'TRIG' fehlt.")
        return df
    except Exception as e: raise ValueError(f"CSV Lesefehler: {str(e)}")

def parse_uploaded_files(uploaded_files):
    participants = {}
    for f in uploaded_files:
        match = re.match(r"([a-zA-Z0-9]+)_([a-zA-Z0-9\-\_]+)\.csv", f.name)
        if match:
            f_type, f_id = match.groups()
            f_type = f_type.lower()
            key = 'export' if 'export' in f_type else 'bl' if 'bl' in f_type else f_type
            if key in ['audio1', 'audio2', 'youtube', 'bl', 'export']:
                if f_id not in participants: participants[f_id] = {}
                participants[f_id][key] = f
    return participants

# --------------------------
# SPEZIAL-METHODEN (ANALYSIS)
# --------------------------

def calculate_physio_ranks(values, threshold):
    """
    Hilfsfunktion: Berechnet 'Physiologische Ränge' (Tied Ranks) 
    basierend auf dem Toleranz-Korridor.
    """
    sort_idx = np.argsort(values)[::-1]
    sorted_values = values[sort_idx]
    tie_eeg_ranks = np.zeros(len(values))
    current_rank_counter = 1
    i = 0
    while i < len(values):
        cluster_indices = [sort_idx[i]]
        j = i + 1
        while j < len(values) and abs(sorted_values[i] - sorted_values[j]) < threshold:
            cluster_indices.append(sort_idx[j])
            j += 1
        mean_rank = current_rank_counter + (len(cluster_indices) - 1) / 2.0
        for idx in cluster_indices: tie_eeg_ranks[idx] = mean_rank
        current_rank_counter += len(cluster_indices)
        i = j
    return tie_eeg_ranks

def calculate_spearman_variants(df_subset, metric_col, rank_col_subj, threshold):
    values = df_subset[metric_col].values
    subj_ranks = df_subset[rank_col_subj].values
    
    # 1. Standard
    std_eeg_ranks = rankdata([-v for v in values], method='average')
    r_std, _ = spearmanr(subj_ranks, std_eeg_ranks)
    
    # 2. Fair / Tie-Break
    tie_eeg_ranks = calculate_physio_ranks(values, threshold)
    r_tie, _ = spearmanr(subj_ranks, tie_eeg_ranks)
    
    # 3. Sehr Fair (Optimiert)
    opt_eeg_ranks = np.zeros(len(values))
    sort_idx = np.argsort(values)[::-1]
    sorted_values = values[sort_idx]
    current_rank_counter = 1
    i = 0
    while i < len(values):
        cluster_indices = [sort_idx[i]]
        j = i + 1
        while j < len(values) and abs(sorted_values[i] - sorted_values[j]) < threshold:
            cluster_indices.append(sort_idx[j])
            j += 1
        available_ranks = list(range(current_rank_counter, current_rank_counter + len(cluster_indices)))
        cluster_subj_ranks = subj_ranks[cluster_indices]
        sorted_rel_indices = np.argsort(cluster_subj_ranks)
        for k, rel_idx in enumerate(sorted_rel_indices):
            opt_eeg_ranks[cluster_indices[rel_idx]] = available_ranks[k]
        current_rank_counter += len(cluster_indices)
        i = j
    r_opt, _ = spearmanr(subj_ranks, opt_eeg_ranks)
    return r_std, r_tie, r_opt

def generate_detailed_stats(df_cases, value_col, agg_func='count'):
    """
    Erstellt eine Statistik-Tabelle, die Grundaktivität in Gesamt, 
    Audio1 und Audio2 aufteilt.
    """
    if agg_func == 'count':
        stats = df_cases['Song'].value_counts().reset_index()
        stats.columns = ['Song', 'Wert']
    else:
        stats = df_cases.groupby('Song')[value_col].mean().reset_index()
        stats.columns = ['Song', 'Wert']
    
    ga_stats = []
    for playlist in ['audio1', 'audio2']:
        mask = (df_cases['Song'] == 'Grundaktivität') & (df_cases['Playlist'] == playlist)
        if mask.any():
            subset = df_cases[mask]
            if agg_func == 'count': val = len(subset)
            else: val = subset[value_col].mean()
            ga_stats.append({
                'Song': f"Grundaktivität ({'Audio 1' if 'audio1' in playlist else 'Audio 2'})",
                'Wert': val
            })
    
    if ga_stats:
        stats = pd.concat([stats, pd.DataFrame(ga_stats)], ignore_index=True)
        
    stats = stats.sort_values(by='Wert', ascending=False).reset_index(drop=True)
    return stats

# Hilfsfunktion für natürliche Sortierung (Song 1, Song 2, ..., Song 10)
def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'([0-9]+)', s)]

# --------------------------
# UI & HELPERS
# --------------------------

def color_tolerance_usage(val):
    if pd.isna(val): return ''
    if val <= 10: return 'background-color: #2ECC71; color: black'
    elif val <= 25: return 'background-color: #82E0AA; color: black'
    elif val <= 50: return 'background-color: #D4EFDF; color: black'
    elif val <= 75: return 'background-color: #F9E79F; color: black'
    elif val <= 90: return 'background-color: #F5B041; color: black'
    else: return 'background-color: #EC7063; color: black'

# --------------------------
# MAIN APP
# --------------------------

def main():
    st.title("🎓 Bachelor EEG Analyse")
    
    if 'master_data' not in st.session_state:
        st.session_state['master_data'] = None

    with st.sidebar:
        st.header("1. Daten Upload")
        uploaded_files = st.file_uploader("CSV Dateien", accept_multiple_files=True)
        cleaning_lvl = st.selectbox("Reinigungs-Level", CLEANING_LEVELS, index=0)
        
        if st.button("Daten einlesen", type="primary"):
            if not uploaded_files:
                st.error("Bitte Dateien hochladen.")
            else:
                participants = parse_uploaded_files(uploaded_files)
                process_data(participants, cleaning_lvl)

    if st.session_state['master_data'] is not None:
        df_full = st.session_state['master_data']
        analysis_df = df_full[df_full['playlist'] != 'youtube'].copy()

        st.markdown("---")
        st.header("2. Analyse Parameter")
        
        with st.container():
            col1, col2 = st.columns(2)
            with col1:
                tol_thresh = st.number_input("Toleranz-Schwellenwert (Physio)", 
                                            min_value=0.01, max_value=2.0, value=0.05, step=0.01,
                                            help="Werte innerhalb dieses Bereichs gelten als physiologisch identisch (Rauschen).")
            with col2:
                high_impact_rank = st.number_input("Min. Rang-Diff für High-Impact", 
                                                  min_value=1, max_value=5, value=2,
                                                  help="Ab welcher Rang-Abweichung wird ein Fehler in Analyse 3 relevant?")

        st.markdown("---")
        
        tabs = st.tabs(["🧠 Alpha/Beta (Wach)", "💤 Theta/Beta (Tief)", "⚖️ (Alpha+Theta)/Beta (Kombi)"])
        
        for i, metric in enumerate(METRICS):
            with tabs[i]:
                st.header(f"Analyse: {metric}")

                # -----------------------------------
                # 1. SPEARMAN
                # -----------------------------------
                st.subheader("1. Rangkorrelation (Spearman)")
                
                with st.expander("📘 Erklärung: Spearman Varianten & Toleranzlogik", expanded=False):
                    st.markdown("### Methodik: Hierarchische Bewertung")
                    st.markdown(
                        "Diese Analyse berücksichtigt die **physiologische Auflösungsgrenze**. Die Grundannahme lautet: "
                        "Das menschliche Gehirn kann extrem feine Unterschiede in den Messwerten (z.B. 0.001) nicht bewusst wahrnehmen. "
                        "Daher unterscheiden wir drei Bewertungsstufen:"
                    )
                    st.markdown(
                        "*   **Standard (Mathematisch streng):** Jede numerische Abweichung führt zu einem unterschiedlichen Rang. Ignoriert Rauschen.\n"
                        "*   **Fair (Tie-Break):** Werte innerhalb des Toleranzbereichs gelten als **physiologisch identisch**. Sie erhalten denselben Durchschnittsrang. Dies bereinigt 'Zufallstreffer', bei denen der Proband eine Reihenfolge geraten hat, die physiologisch nicht signifikant ist.\n"
                        "*   **Sehr Fair (Subjektiv optimiert):** Innerhalb des Toleranzbereichs (physiologische Indifferenz) wird die **subjektive Wahrnehmung als 'Tie-Breaker'** akzeptiert. Der Algorithmus nimmt an: Wenn das Gehirn keinen Unterschied macht, hat das Gefühl des Probanden recht."
                    )
                    st.markdown("---")
                    st.markdown("### ⚙️ Technische Umsetzung")
                    st.markdown(
                        r"1.  **Delta-Clustering:** Der Algorithmus sortiert die EEG-Werte und gruppiert benachbarte Werte, deren Differenz $\Delta < Schwellenwert$ ist." + "\n"
                        "2.  **Lokale Rang-Permutation:**\n"
                        "    *   *Fair:* Alle Items im Cluster erhalten den arithmetischen Mittelwert der verfügbaren Ränge.\n"
                        "    *   *Sehr Fair:* Die EEG-Ränge innerhalb des Clusters werden so permutiert (umgestellt), dass sie die Rangkorrelation zum subjektiven Ranking des Probanden lokal maximieren."
                    )
                    st.info(
                        "💡 **Hinweis:** Sinken die Werte bei 'Fair'/'Sehr Fair', bedeutet das oft, dass der Proband "
                        "eine 'richtige' Unterscheidung getroffen hat, obwohl die Werte im Rauschbereich lagen (Lucky Guess), "
                        "oder dass die lokale Optimierung im Widerspruch zum globalen Trend steht."
                    )

                spearman_res = []
                if not analysis_df.empty:
                    for pid in analysis_df['participant'].unique():
                        p_data = analysis_df[analysis_df['participant'] == pid]
                        r_std, r_tie, r_opt = calculate_spearman_variants(p_data, metric, 'subj_rank', tol_thresh)
                        spearman_res.append({
                            'Proband': pid, 'Standard': r_std, 'Fair (Tie)': r_tie, 'Sehr Fair (Opt)': r_opt,
                            'Verbesserung': f"{r_tie - r_std:+.3f} ; {r_opt - r_std:+.3f}"
                        })
                    st.dataframe(pd.DataFrame(spearman_res).style.format("{:.3f}", subset=['Standard', 'Fair (Tie)', 'Sehr Fair (Opt)']).background_gradient(cmap='RdYlGn', subset=['Standard', 'Fair (Tie)', 'Sehr Fair (Opt)'], vmin=-1, vmax=1))
                else: st.warning("Keine Daten.")

                # -----------------------------------
                # 2. TOLERANZ KORRIDOR
                # -----------------------------------
                st.subheader("2. Der „Toleranz-Korridor“ (Logic Check)")
                
                with st.expander("📘 Erklärung: Toleranz-Ausnutzung & Nähe", expanded=False):
                    st.markdown("### Methodik: Detektion physiologischer „Zwillinge“")
                    st.markdown(
                        "Dieser Logic-Check identifiziert Song-Paare, die für den Probanden physiologisch kaum unterscheidbar waren. "
                        "Er dient als Validierung dafür, ob eine „falsche“ Reihenfolge im Ranking überhaupt relevant ist."
                    )
                    st.markdown(
                        "*   **Physiologisches Cluster:** Liegt die Differenz zweier Songs unter dem Schwellenwert, betrachtet das Gehirn diese als denselben Zustand.\n"
                        "*   **Ausnutzung (%):** Gibt an, wie nah die Differenz am Grenzwert liegt. Eine geringe Prozentzahl deutet auf extrem hohe Ähnlichkeit hin."
                    )
                    st.markdown("---")
                    st.markdown("### ⚙️ Technische Umsetzung")
                    st.markdown(
                        r"1.  **Differenz-Berechnung:** Für jedes benachbarte Song-Paar (nach EEG-Wert sortiert) wird $\Delta = |Val_A - Val_B|$ berechnet." + "\n"
                        "2.  **Visualisierung (Ampel-Logik):**\n"
                        "    *   🟢 **< 25%:** Werte sind nahezu identisch (Messrauschen).\n"
                        "    *   🟡 **25% - 75%:** Werte sind ähnlich, aber unterscheidbar.\n"
                        "    *   🔴 **> 90%:** Werte kratzen an der Grenze zur Signifikanz."
                    )
                    st.info(
                        "💡 **Interpretation:** Tauchen hier viele Paare auf, deutet das darauf hin, dass die Playlist für den Probanden "
                        "emotional/kognitiv sehr homogen war. Ranking-Fehler sind in diesem Fall erwartbar."
                    )

                tol_cases = []
                if not analysis_df.empty:
                    for _, grp in analysis_df.groupby(['participant', 'playlist']):
                        grp = grp.sort_values(metric, ascending=False)
                        vals = grp[metric].values; names = grp['item'].values
                        for k in range(len(vals)-1):
                            diff = abs(vals[k] - vals[k+1])
                            if diff < tol_thresh:
                                tol_cases.append({
                                    'Proband': grp['participant'].iloc[0], 'Playlist': grp['playlist'].iloc[0],
                                    'Items': f"{names[k]} ↔ {names[k+1]}", 'Diff': diff,
                                    'Ausnutzung (%)': (diff / tol_thresh) * 100
                                })
                if tol_cases: st.dataframe(pd.DataFrame(tol_cases).style.format({'Diff': '{:.4f}', 'Ausnutzung (%)': '{:.1f}%'}).map(color_tolerance_usage, subset=['Ausnutzung (%)']))
                else: st.info(f"Keine Fälle < {tol_thresh}.")

                # -----------------------------------
                # 3. HIGH IMPACT (STANDARD)
                # -----------------------------------
                st.subheader("3. High-Impact Error (Standard Rang-Differenz)")
                
                with st.expander("📘 Erklärung: Kritische Rang-Divergenz", expanded=False):
                    st.markdown("### Methodik: Identifikation grober Fehleinschätzungen")
                    st.markdown(
                        "Diese Analyse filtert Fälle, in denen das subjektive Empfinden (User-Ranking) stark von der objektiven Messung (EEG-Ranking) abweicht. "
                        "Es handelt sich um einen rein mathematischen Filter, der das physiologische Rauschen noch **nicht** berücksichtigt."
                    )
                    st.markdown(
                        "*   **Divergenz-Kriterium:** Ein Fehler gilt als „High Impact“, wenn die Differenz zwischen subjektivem Rang und EEG-Rang den eingestellten Grenzwert überschreitet.\n"
                        "*   **Zweck:** Dient als „Brutto-Fehlerliste“ vor der physiologischen Bereinigung in Schritt 4."
                    )
                    st.markdown("---")
                    st.markdown("### ⚙️ Technische Umsetzung")
                    st.markdown(
                        r"1.  **Ranking-Vergleich:** $Delta_{Rang} = |Rang_{Subj} - Rang_{EEG}|$" + "\n"
                        r"2.  **Filterung:** Zeige Datensatz, wenn $Delta_{Rang} \ge HighImpactRank$."
                    )
                    st.info(
                        "💡 **Hinweis:** Ein hier angezeigter Fehler muss nicht zwingend physiologisch signifikant sein. "
                        "Vergleichen Sie die Ergebnisse immer mit Analyse 4, um reine „Knapp-daneben-Treffer“ (Toleranzbereich) auszuschließen."
                    )

                impact_cases = []
                if not analysis_df.empty:
                    for idx, row in analysis_df.iterrows():
                        eeg_r = row[f'rank_eeg_{metric}']
                        sub_r = row['subj_rank']
                        if abs(eeg_r - sub_r) >= high_impact_rank:
                            impact_cases.append({
                                'Proband': row['participant'], 'Playlist': row['playlist'],
                                'Song': row['item'], 'Rang-Diff': int(abs(eeg_r - sub_r))
                            })
                
                if impact_cases:
                    imp_df = pd.DataFrame(impact_cases)
                    c1, c2 = st.columns([2, 1])
                    with c1: 
                        st.markdown("**Gefundene Fehler:**")
                        st.dataframe(imp_df)
                    with c2: 
                        st.markdown("**Statistik (Anzahl Fehler):**")
                        stats_std = generate_detailed_stats(imp_df, 'Rang-Diff', agg_func='count')
                        st.dataframe(stats_std)
                else: st.success("Keine Standard-Fehler gefunden.")

                # -----------------------------------
                # 4. WEIGHTED HIGH IMPACT
                # -----------------------------------
                st.subheader("4. Physiologisch korrigierte High-Impact Analyse")
                
                with st.expander("📘 Erklärung: Hybrid-Modell & Weighted Distance Score", expanded=True):
                    st.error(f"⚠️ **Zentrale Abhängigkeit:** Diese Analyse basiert vollständig auf dem Parameter **Toleranz-Schwellenwert (Aktuell: {tol_thresh})**.")
                    
                    st.markdown("### Methodik: Rauschbereinigte Fehleranalyse")
                    st.markdown(
                        "Dieses Hybrid-Modell kombiniert statistische Fairness mit physiologischer Signalstärke. "
                        "Ziel ist es, **reines Messrauschen zu ignorieren** und nur solche Fehler hervorzuheben, die auf einer echten Diskrepanz zwischen Wahrnehmung und Physiologie beruhen."
                    )
                    st.markdown("---")
                    st.markdown("### ⚙️ Technische Umsetzung (Der 3-Stufen-Filter)")
                    st.markdown(
                        "1.  **Stufe 1: Tied Ranks (Bereinigung):**\n"
                        "    Die EEG-Ränge werden neu berechnet. Werte, die sich um weniger als den Toleranz-Schwellenwert unterscheiden, erhalten **denselben Rang** (Cluster-Bildung).\n"
                        "2.  **Stufe 2: Der „Gatekeeper“ (Harter Filter):**\n"
                        "    Wir vergleichen den gemessenen Wert ($Value_{Actual}$) mit dem theoretischen Referenzwert ($Value_{Ref}$), den der Song auf dem vom User gewählten Rang haben müsste.\n"
                        "    *   Regel: Ist $|Value_{Actual} - Value_{Ref}| < Toleranz$, wird der **Score auf 0 gesetzt**.\n"
                        "3.  **Stufe 3: Weighted Score Berechnung:**\n"
                        "    Nur wenn der Filter passiert wird, berechnet sich der Score:\n"
                        r"    $$ Score = |Rang_{Subj} - Rang_{Tied}| \times |Value_{Actual} - Value_{Ref}| $$"
                    )
                    st.info(
                        "💡 **Interpretation:**\n"
                        "*   **Einzelfehler (Links):** Zeigt nur Fehler mit Score > 0. Verschwindet ein Song hier im Vergleich zu Sektion 3, war der Fehler physiologisch irrelevantes Rauschen.\n"
                        "*   **Gesamtstatistik (Rechts):** Der Durchschnittswert (Mean) berücksichtigt **alle** Songs. Ein hoher Wert signalisiert Songs, bei denen sich Probanden systematisch und signifikant irren."
                    )

                weighted_cases = []
                if not analysis_df.empty:
                    for (pid, playlist), group in analysis_df.groupby(['participant', 'playlist']):
                        vals = group[metric].values
                        tied_ranks = calculate_physio_ranks(vals, tol_thresh)
                        sorted_vals_desc = np.sort(vals)[::-1]
                        
                        temp_grp = group.copy(); temp_grp['tied_rank'] = tied_ranks
                        
                        for idx, row in temp_grp.iterrows():
                            s_rank = int(row['subj_rank'])
                            t_rank = row['tied_rank']
                            val_act = row[metric]
                            ref_idx = min(s_rank - 1, len(sorted_vals_desc) - 1)
                            val_ref = sorted_vals_desc[ref_idx]
                            
                            rank_diff_tied = abs(s_rank - t_rank)
                            val_diff = abs(val_act - val_ref)
                            
                            is_physio_error = val_diff >= tol_thresh
                            weighted_score = (rank_diff_tied * val_diff) if is_physio_error else 0.0
                            
                            weighted_cases.append({
                                'Proband': row['participant'], 'Playlist': row['playlist'],
                                'Song': row['item'], 'Rang-Diff (Tied)': round(rank_diff_tied, 2),
                                'Δ Messwert (Real vs. Ref)': round(val_diff, 4),
                                'Weighted Score': round(weighted_score, 4)
                            })

                if weighted_cases:
                    w_df = pd.DataFrame(weighted_cases)
                    relevant_errors = w_df[w_df['Weighted Score'] > 0].sort_values(by='Weighted Score', ascending=False)
                    
                    c1, c2 = st.columns([2, 1])
                    with c1:
                        st.markdown("**Signifikante Einzelfehler (Score > 0):**")
                        if not relevant_errors.empty:
                            st.dataframe(relevant_errors.style.background_gradient(subset=['Weighted Score'], cmap='Reds'))
                        else:
                            st.info("Alle Rang-Fehler liegen innerhalb der physiologischen Toleranz (Score = 0).")
                    with c2:
                        st.markdown("**Gesamtübersicht (Ø Weighted Score):**")
                        stats_weighted = generate_detailed_stats(w_df, 'Weighted Score', agg_func='mean')
                        st.dataframe(stats_weighted.style.background_gradient(subset=['Wert'], cmap='Reds').format({'Wert': '{:.4f}'}))

                else: st.success("Keine Daten für Berechnung.")

        # -----------------------------------
        # 5. GLOBALE META-ANALYSE
        # -----------------------------------
        st.markdown("---")
        st.header("5. Globale Aussagen (Meta-Analyse)")
        
        st.info(
            "Dieser Bereich aggregiert die Daten aus allen drei vorherigen Tabs. "
            "Ziel ist es, Trends zu erkennen, die unabhängig von einer einzelnen Berechnungsmethode Bestand haben."
        )

        meta_correlations = []
        meta_weighted_scores = []
        
        for m in METRICS:
            # A) KORRELATIONEN SAMMELN (Für 5.1)
            if not analysis_df.empty:
                r_vals_fair = []
                r_vals_opt = []
                
                for pid in analysis_df['participant'].unique():
                    p_data = analysis_df[analysis_df['participant'] == pid]
                    _, r_tie, r_opt = calculate_spearman_variants(p_data, m, 'subj_rank', tol_thresh)
                    if not np.isnan(r_tie): r_vals_fair.append(r_tie)
                    if not np.isnan(r_opt): r_vals_opt.append(r_opt)
                
                if r_vals_fair:
                    meta_correlations.append({
                        'Metrik': m,
                        'Ø Korrelation (Fair)': np.mean(r_vals_fair),
                        'Ø Korrelation (Sehr Fair)': np.mean(r_vals_opt)
                    })

            # B) WEIGHTED SCORES SAMMELN (Für 5.2)
            for (pid, playlist), group in analysis_df.groupby(['participant', 'playlist']):
                vals = group[m].values
                tied_ranks = calculate_physio_ranks(vals, tol_thresh)
                sorted_vals_desc = np.sort(vals)[::-1]
                temp_grp = group.copy(); temp_grp['tied_rank'] = tied_ranks
                
                for idx, row in temp_grp.iterrows():
                    s_rank = int(row['subj_rank'])
                    t_rank = row['tied_rank']
                    val_act = row[m]
                    ref_idx = min(s_rank - 1, len(sorted_vals_desc) - 1)
                    val_ref = sorted_vals_desc[ref_idx]
                    rank_diff_tied = abs(s_rank - t_rank)
                    val_diff = abs(val_act - val_ref)
                    is_physio_error = val_diff >= tol_thresh
                    w_score = (rank_diff_tied * val_diff) if is_physio_error else 0.0
                    
                    item_name = row['item']
                    specific_name = item_name
                    if item_name == 'Grundaktivität':
                        suffix = 'Audio 1' if 'audio1' in row['playlist'] else 'Audio 2'
                        specific_name = f"Grundaktivität ({suffix})"
                        meta_weighted_scores.append({'Metrik': m, 'Song': 'Grundaktivität (Gesamt)', 'Weighted Score': w_score})
                    
                    meta_weighted_scores.append({'Metrik': m, 'Song': specific_name, 'Weighted Score': w_score})

        # -----------------------------------
        # 5.1 LEADERBOARD
        # -----------------------------------
        st.subheader("5.1 Dominierender Entspannungstypus (Metriken-Vergleich)")
        
        with st.expander("📘 Erklärung: Interpretation der Korrelationswerte", expanded=False):
            st.markdown("### Was bedeuten die Zahlen?")
            st.markdown("Die Werte basieren auf dem **Spearman-Rangkorrelationskoeffizienten** ($r_s$).")
            st.markdown("""
            | Wert ($r_s$) | Interpretation | Bedeutung für die Analyse |
            | :--- | :--- | :--- |
            | **0.5 bis 1.0** | **Starker Zusammenhang** | ✅ Das subjektive Empfinden stimmt stark mit der physiologischen Messung überein. (Grün) |
            | **0.3 bis 0.5** | **Mäßiger Zusammenhang** | ⚠️ Es gibt eine Tendenz, aber viele individuelle Abweichungen. (Gelb/Hellgrün) |
            | **-0.3 bis 0.3** | **Kein Zusammenhang** | ❌ Das Ranking gleicht einem Zufallsergebnis. (Orange/Gelb) |
            | **< -0.3** | **Gegenläufig** | ❌ Probanden empfanden Entspannung genau gegenteilig zur Messung. (Rot) |
            """)
            st.markdown("### Fair vs. Sehr Fair")
            st.markdown("* **Ø Fair:** Wissenschaftlich ehrliches Maß (mit Toleranz).\n* **Ø Sehr Fair:** Optimiertes „Best-Case“ Szenario.")

        if meta_correlations:
            df_meta_corr = pd.DataFrame(meta_correlations).set_index('Metrik')
            df_meta_corr = df_meta_corr.sort_values(by='Ø Korrelation (Fair)', ascending=False)
            st.table(df_meta_corr.style.background_gradient(cmap='RdYlGn', vmin=-1, vmax=1).format("{:.3f}"))
            
            best_metric = df_meta_corr.index[0]
            if df_meta_corr.iloc[0]['Ø Korrelation (Fair)'] > 0.3:
                st.success(f"📌 **Ergebnis:** Die Metrik **„{best_metric}“** zeigt die höchste Übereinstimmung.")
            else:
                st.warning(f"📌 **Ergebnis:** **„{best_metric}“** führt, aber die Korrelation ist insgesamt schwach (< 0.3).")
        
        # -----------------------------------
        # 5.2 ITEM DIFFICULTY
        # -----------------------------------
        st.subheader("5.2 Globale Einschätzungs-Schwierigkeit (Item-Analyse)")
        
        with st.expander("📘 Erklärung: Globale Fehler-Integrität", expanded=False):
            st.markdown("**Fragestellung:** Welche Zustände waren universell schwer einzuschätzen?\n**Berechnung:** Ø Weighted Distance Score über alle Metriken.")

        if meta_weighted_scores:
            df_meta_ws = pd.DataFrame(meta_weighted_scores)
            item_difficulty = df_meta_ws.groupby('Song')['Weighted Score'].mean().reset_index()
            item_difficulty.columns = ['Song / Zustand', 'Ø Globaler Fehler-Score']
            item_difficulty = item_difficulty.sort_values(by='Ø Globaler Fehler-Score', ascending=False).reset_index(drop=True)
            
            c1, c2 = st.columns([2, 1])
            with c1: st.dataframe(item_difficulty.style.background_gradient(subset=['Ø Globaler Fehler-Score'], cmap='Reds').format({'Ø Globaler Fehler-Score': '{:.4f}'}))
            with c2:
                if not item_difficulty.empty:
                    st.markdown(f"**Analyse-Highlights:**\n🔴 **Schwierig:** *{item_difficulty.iloc[0]['Song / Zustand']}*\n🟢 **Klar:** *{item_difficulty.iloc[-1]['Song / Zustand']}*")

        # -----------------------------------
        # 6. GENRE ANALYSE (LEADERBOARD)
        # -----------------------------------
        st.markdown("---")
        st.header("6. Genre-Analyse: Die „Entspannungs-Meisterschaft“")
        
        st.info(
            "Dieser Bereich vergleicht, welches Musik-Genre **subjektiv** (nach Meinung) vs. **objektiv** (nach EEG) "
            "am besten abgeschnitten hat. Durch ein **intra-subjektives Ranking** werden Unterschiede in der Signalstärke (z.B. dicker Schädel) herausgerechnet."
        )

        # --- 6.1 MAPPING ---
        st.subheader("6.1 Genre-Definition (Mapping)")
        
        # Songs sammeln und natürlich sortieren (1, 2, ... 10)
        raw_unique_songs = analysis_df[~analysis_df['item'].str.contains('Grundaktivität', case=False, na=False)]['item'].unique()
        # Sortieren mit natural_sort_key
        unique_songs = sorted(raw_unique_songs, key=natural_sort_key)
        
        genre_mapping = {}
        
        with st.expander("🎵 Songs zu Genres zuordnen (Bitte ausfüllen)", expanded=True):
            # Layout Optimierung: 4 Spalten für bessere Übersicht
            cols = st.columns(4)
            for i, song in enumerate(unique_songs):
                with cols[i % 4]:
                    # Default Wert aus Config laden, falls vorhanden
                    default_val = DEFAULT_GENRES.get(song, "")
                    val = st.text_input(f"{song}", value=default_val, placeholder="Genre...", key=f"genre_{i}")
                    genre_mapping[song] = val.strip().title() if val.strip() != "" else "Unbekannt"
        
        # Daten vorbereiten
        genre_df = analysis_df.copy()
        genre_df['genre'] = genre_df['item'].apply(lambda x: 'Ruhephase' if 'Grundaktivität' in x else genre_mapping.get(x, 'Unbekannt'))
        
        # --- DATEN AGGREGATION ---
        
        # Auswahl der Metrik für das EEG-Leaderboard
        target_metric_genre = st.selectbox(
            "Wählen Sie die physiologische Metrik für das EEG-Leaderboard:", 
            METRICS, index=2, 
            help="Welcher physiologische Zustand soll als 'Sieger' definiert werden?"
        )

        # 1. Gruppieren pro Proband & Genre (Mittelwert bilden, falls Genre mehrfach vorkommt)
        grp_genre = genre_df.groupby(['participant', 'genre']).agg({
            'subj_rank': 'mean',
            target_metric_genre: 'mean'
        }).reset_index()

        # 2. Ranking pro Proband erstellen (Normalisierung)
        # Subjektiv: 1 (klein) ist gut
        grp_genre['rank_norm_subj'] = grp_genre.groupby('participant')['subj_rank'].rank(method='average', ascending=True)
        # EEG: Hoher Wert ist gut -> ascending=False
        grp_genre['rank_norm_eeg'] = grp_genre.groupby('participant')[target_metric_genre].rank(method='average', ascending=False)
        
        # 3. Globale Durchschnitte der Ränge berechnen
        leaderboard = grp_genre.groupby('genre').agg({
            'rank_norm_subj': 'mean',
            'rank_norm_eeg': 'mean'
        }).reset_index()
        
        leaderboard.columns = ['Genre', 'Ø Rang (Subjektiv)', 'Ø Rang (EEG)']
        
        # Sortieren für Visualisierung
        leaderboard_subj = leaderboard.sort_values('Ø Rang (Subjektiv)')
        leaderboard_eeg = leaderboard.sort_values('Ø Rang (EEG)')
        
        # --- 6.2 ANALYSE A: SUBJEKTIV ---
        st.subheader("6.2 Analyse A: Die Erwartung (Subjektives Ranking)")
        
        with st.expander("📘 Erklärung: Das Glaubens-Ranking", expanded=False):
            st.markdown(
                "**Fragestellung:** Welches Genre *glaubten* die Probanden, sei am entspannendsten?\n"
                "**Metrik:** Durchschnittlicher Rangplatz (1 = Bester)."
            )
            
        c1, c2 = st.columns([1, 1])
        with c1:
            st.markdown("**Leaderboard (Gefühl):**")
            # KORREKTUR: subset hinzufügen, damit 'Genre' (Text) nicht formatiert wird
            st.dataframe(
                leaderboard_subj[['Genre', 'Ø Rang (Subjektiv)']]
                .style
                .background_gradient(cmap='Greens_r', subset=['Ø Rang (Subjektiv)'])
                .format("{:.2f}", subset=['Ø Rang (Subjektiv)']) 
            )
        with c2:
            chart_s = alt.Chart(leaderboard_subj).mark_bar().encode(
                x=alt.X('Ø Rang (Subjektiv)', title='Ø Platzierung (1=Top)'),
                y=alt.Y('Genre', sort='x'),
                color=alt.Color('Ø Rang (Subjektiv)', scale=alt.Scale(scheme='greens', reverse=True), legend=None)
            )
            st.altair_chart(chart_s, width='stretch')

        # --- 6.3 ANALYSE B: OBJEKTIV ---
        st.subheader("6.3 Analyse B: Die physiologische Realität (EEG Ranking)")
        
        with st.expander(f"📘 Erklärung: Das EEG-Ranking (Basis: {target_metric_genre})", expanded=False):
            st.markdown(
                "**Fragestellung:** Bei welchem Genre zeigte das Gehirn *tatsächlich* die stärkste Entspannungsreaktion?\n"
                "**Normalisierung:** Da EEG-Amplituden individuell sind, zählt nur der **relative Rang** pro Proband."
            )
            
        c1, c2 = st.columns([1, 1])
        with c1:
            st.markdown("**Leaderboard (Messung):**")
            st.dataframe(
                leaderboard_eeg[['Genre', 'Ø Rang (EEG)']]
                .style
                .background_gradient(cmap='Blues_r', subset=['Ø Rang (EEG)'])
                .format("{:.2f}", subset=['Ø Rang (EEG)'])
            )
        with c2:
            chart_e = alt.Chart(leaderboard_eeg).mark_bar().encode(
                x=alt.X('Ø Rang (EEG)', title='Ø Platzierung (1=Top)'),
                y=alt.Y('Genre', sort='x'),
                color=alt.Color('Ø Rang (EEG)', scale=alt.Scale(scheme='blues', reverse=True), legend=None)
            )
            st.altair_chart(chart_e, width='stretch')

        # --- FAZIT ---
        st.markdown("---")
        st.markdown("### 🔍 Vergleich: Wahrnehmung vs. Realität")
        
        # Delta berechnen mit korrekten Spaltennamen
        leaderboard['Delta (Subj - EEG)'] = leaderboard['Ø Rang (Subjektiv)'] - leaderboard['Ø Rang (EEG)']
        leaderboard['Interpretation'] = leaderboard['Delta (Subj - EEG)'].apply(
            lambda x: "Physiologisch besser als gedacht" if x > 0.5 else ("Physiologisch schlechter als gedacht" if x < -0.5 else "Wahrnehmung korrekt ✅")
        )
        
        st.dataframe(
            leaderboard.style
            .background_gradient(subset=['Delta (Subj - EEG)'], cmap='RdYlGn', vmin=-2, vmax=2)
            .format("{:.2f}", subset=['Ø Rang (Subjektiv)', 'Ø Rang (EEG)', 'Delta (Subj - EEG)'])
        )

# --------------------------
# DATEN VERARBEITUNG
# --------------------------
def process_data(participants, cleaning_lvl):
    master_data = []
    progress = st.progress(0)
    for idx, (p_id, files) in enumerate(participants.items()):
        if 'export' not in files or 'bl' not in files: continue
        try:
            ranking_df = pd.read_csv(files['export']); ranking_df.columns = ranking_df.columns.str.strip()
            bl_df = load_eeg_file(files['bl'], "BL")
            bl_ev = get_global_events(bl_df)
            bl_pwr = compute_band_power(clean_segment(bl_df, cleaning_lvl, 2.5, bl_ev), SAMPLING_FREQUENCY)
            if not bl_pwr: continue
            bl_ratios = calculate_ratios(bl_pwr)
            
            for pl_key in ['audio1', 'audio2', 'youtube']:
                if pl_key not in files: continue
                pl_df = load_eeg_file(files[pl_key], pl_key)
                pl_ev = get_global_events(pl_df)
                segments = segment_song_data(pl_df)
                r_map = {'audio1': 'audio_1', 'audio2': 'audio_2', 'youtube': 'youtube'}
                subset = ranking_df[ranking_df['ranking_type'] == r_map[pl_key]]
                if subset.empty: continue
                
                try: ga_rank = subset[subset['item_name'] == 'Grundaktivität']['final_rank'].min()
                except: ga_rank = 6
                song_names = subset[~subset['item_name'].isin(['Grundaktivität', 'baseline'])].sort_values('play_order')['item_name'].tolist()
                
                pl_items = []
                pl_items.append({
                    'participant': p_id, 'playlist': pl_key, 'item': 'Grundaktivität',
                    'is_baseline': True, 'subj_rank': ga_rank, **bl_ratios
                })
                for i, seg in enumerate(segments):
                    if i >= len(song_names): break
                    s_pwr = compute_band_power(clean_segment(seg, cleaning_lvl, 2.5, pl_ev), SAMPLING_FREQUENCY)
                    if s_pwr:
                        s_ratios = calculate_ratios(s_pwr)
                        s_rank = subset[subset['item_name'] == song_names[i]]['final_rank'].min()
                        pl_items.append({
                            'participant': p_id, 'playlist': pl_key, 'item': song_names[i],
                            'is_baseline': False, 'subj_rank': s_rank, **s_ratios
                        })
                df_pl = pd.DataFrame(pl_items)
                for m in METRICS: df_pl[f'rank_eeg_{m}'] = df_pl[m].rank(ascending=False)
                master_data.append(df_pl)
        except Exception: pass
        progress.progress((idx + 1) / len(participants))
    
    if master_data:
        st.session_state['master_data'] = pd.concat(master_data)
        st.success("Verarbeitung abgeschlossen!"); st.rerun()
    else: st.error("Keine validen Daten generiert.")

if __name__ == "__main__":
    main()