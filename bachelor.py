import streamlit as st
import pandas as pd
import numpy as np
from scipy.signal import welch
from scipy.stats import spearmanr, rankdata, kendalltau
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
# Die Namen (Keys) müssen mit den Songnamen in der CSV übereinstimmen.
DEFAULT_GENRES = {
    "Song 1": "Klassik",
    "Song 2": "Electronic [Ambient]",
    "Song 3": "Klassik",
    "Song 4": "Electronic [Ambient]",
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

def clean_segment(segment, cleaning_level, short_radius_s, reaction_buffer_s, global_events):
    """
    Bereinigt ein EEG-Segment basierend auf Level, Radius und Reaktions-Puffer.
    """
    if cleaning_level == 'Ohne' or segment.empty: return segment
    
    longs, trans, shorts = global_events
    indices_to_drop = set()
    s_start, s_end = segment.index.min(), segment.index.max()
    
    # Umrechnung Sekunden in Samples (Zeilen)
    buffer_rows = int(reaction_buffer_s * SAMPLING_FREQUENCY)
    radius_rows = int(short_radius_s * SAMPLING_FREQUENCY)

    def overlap(ev_start, ev_end): 
        return max(ev_start, s_start) <= min(ev_end, s_end)

    if cleaning_level in ['Normal', 'Strikt', 'Sehr Strikt']:
        for t_s, t_e in trans:
            if overlap(t_s, t_e):
                drop_start = max(t_s, s_start)
                drop_end = min(t_e, s_end)
                indices_to_drop.update(segment.loc[drop_start:drop_end].index)

    if cleaning_level in ['Strikt', 'Sehr Strikt']:
        for l_s, l_e in longs:
            # Puffer anwenden: Startpunkt liegt früher
            adjusted_start = l_s - buffer_rows
            if overlap(adjusted_start, l_e): 
                drop_start = max(adjusted_start, s_start)
                drop_end = min(l_e, s_end)
                indices_to_drop.update(segment.loc[drop_start:drop_end].index)

    if cleaning_level == 'Sehr Strikt':
        for sh_idx in shorts:
            if s_start <= sh_idx <= s_end:
                # Links: Radius + Puffer | Rechts: Nur Radius
                d_start = max(s_start, sh_idx - (radius_rows + buffer_rows))
                d_end = min(s_end, sh_idx + radius_rows)
                indices_to_drop.update(segment.loc[d_start:d_end].index)
                
    return segment.drop(index=list(indices_to_drop), errors='ignore')

def segment_song_data(df):
    # 1. Spalte prüfen
    if 'TRIG' not in df.columns: return []
    
    # 2. Alle relevanten Trigger (1=Start, 2=Ende) isolieren
    # Wir brauchen die Indices, um zu wissen, wo wir schneiden müssen. Cleaning passiert danach.
    all_triggers = df[df['TRIG'].isin([1, 2])]
    if all_triggers.empty: return []
    
    starts = all_triggers[all_triggers['TRIG'] == 1].index.tolist()
    ends = all_triggers[all_triggers['TRIG'] == 2].index.tolist()
    
    if not starts: return []
    
    # 3. Master End Punkt bestimmen
    # Standardmäßig ist das Ende der Datei das Limit (für den Fall 1->1->EOF)
    master_end = df.index[-1]
    
    # Nur wenn wir End-Trigger haben UND der letzte End-Trigger tatsächlich
    # NACH dem letzten Start-Trigger kommt, gilt dieser als hartes Ende der Playlist.
    if ends and starts and ends[-1] > starts[-1]:
        master_end = ends[-1]
        
    # 4. Kombinierte Liste aller Schnittpunkte erstellen
    # Wir nehmen alle Starts und Ends, die zeitlich vor oder genau auf dem Master-End liegen
    trigger_points = sorted([idx for idx in starts + ends if idx <= master_end])
    
    segments = []
    
    for start_idx in starts:
        # Falls ein Start-Trigger fälschlicherweise nach dem Master-End liegt, ignorieren
        if start_idx >= master_end: continue
        
        # Das Ende für DIESEN Song bestimmen
        current_end = master_end
        
        # Wir suchen den nächstgelegenen Trigger (egal ob 1 oder 2), der nach dem Start kommt.
        # Das löst das 1 -> 1 -> 1 Problem, da der nächste "1er" als Ende für den aktuellen gilt.
        for trig_idx in trigger_points:
            if trig_idx > start_idx:
                current_end = trig_idx
                break
        
        # Segment ausschneiden
        # Hinweis: .loc ist inklusive des End-Index.
        segments.append(df.loc[start_idx:current_end])
        
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
        # 1. Bereinigung des Dateinamens (Shuffle und Merged für Export(Ranking) entfernen), da Zuordnung der Files via filename passiert.
        is_shuffle_file = 'shuffle' in f.name.lower()
        
        # HIER IST DIE ÄNDERUNG: .replace('_merged', '') hinzugefügt
        clean_name = f.name.replace('_shuffle', '').replace('-shuffle', '').replace('_merged', '')
        
        # 2. Regex für Format: Typ_ID.csv
        match = re.match(r"([a-zA-Z0-9]+)_([a-zA-Z0-9\-\_]+)\.csv", clean_name)
        
        if match:
            f_type_raw, f_id = match.groups()
            f_type = f_type_raw.lower()
            
            # 3. Logik zur Bestimmung des internen keys
            
            # Fall A: Baseline
            if 'bl' in f_type:
                key = 'bl_shuffle' if is_shuffle_file else 'bl'
            
            # Fall B: Youtube
            elif f_type == 'youtube':
                key = 'youtube_shuffle' if is_shuffle_file else 'youtube'
            
            # Fall C: Export
            elif f_type == 'export':
                key = 'export'
            
            # Fall D: Audio Files
            else:
                key = f_type 

            # 4. Speichern
            valid_keys = [
                'audio1', 'audio2', 'audio3', 'audio4', 
                'youtube', 'youtube_shuffle', 
                'bl', 'bl_shuffle', 'export'
            ]
            
            if key in valid_keys:
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

def calculate_correlation_variants(df_subset, metric_col, rank_col_subj, threshold):
    """
    Berechnet Spearman UND Kendall für Standard, Fair und Sehr Fair.
    Gibt zwei Tupel zurück: (s_std, s_tie, s_opt), (k_std, k_tie, k_opt)
    """
    values = df_subset[metric_col].values
    subj_ranks = df_subset[rank_col_subj].values
    
    # --- 1. Standard ---
    std_eeg_ranks = rankdata([-v for v in values], method='average')
    s_std = spearmanr(subj_ranks, std_eeg_ranks)[0]
    k_std = kendalltau(subj_ranks, std_eeg_ranks)[0]
    
    # --- 2. Fair / Tie-Break ---
    tie_eeg_ranks = calculate_physio_ranks(values, threshold)
    s_tie = spearmanr(subj_ranks, tie_eeg_ranks)[0]
    k_tie = kendalltau(subj_ranks, tie_eeg_ranks)[0]
    
    # --- 3. Sehr Fair (Optimiert) ---
    # Wir nutzen den Permutations-Algorithmus, der für Spearman optimiert wurde,
    # wenden darauf aber auch Kendall an (da die Rangfolge identisch ist).
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
        for k_idx, rel_idx in enumerate(sorted_rel_indices):
            opt_eeg_ranks[cluster_indices[rel_idx]] = available_ranks[k_idx]
        current_rank_counter += len(cluster_indices)
        i = j
    
    s_opt = spearmanr(subj_ranks, opt_eeg_ranks)[0]
    k_opt = kendalltau(subj_ranks, opt_eeg_ranks)[0]
    
    return (s_std, s_tie, s_opt), (k_std, k_tie, k_opt)

def generate_detailed_stats(df_cases, value_col, agg_func='count'):
    """
    Erstellt eine Statistik-Tabelle, die Grundaktivität in Gesamt
    sowie Audio 1, 2, 3, 4 aufteilt.
    """
    # 1. Basis-Statistik (Hier ist "Grundaktivität" = GESAMT aller Playlists)
    if agg_func == 'count':
        stats = df_cases['Song'].value_counts().reset_index()
        stats.columns = ['Song', 'Wert']
    else:
        stats = df_cases.groupby('Song')[value_col].mean().reset_index()
        stats.columns = ['Song', 'Wert']
    
    # Den allgemeinen Eintrag explizit als "(Gesamt)" kennzeichnen
    stats['Song'] = stats['Song'].replace('Grundaktivität', 'Grundaktivität (Gesamt)')
    
    # 2. Spezifische Statistik pro Playlist hinzufügen (shuffle ist Indikator für zweite Aufnahmesession)
    ga_stats = []
    target_playlists = ['audio1', 'audio2', 'audio3', 'audio4', 'youtube', 'youtube_shuffle']
    
    for playlist_key in target_playlists:
        # Filter: Ist Grundaktivität UND gehört zur Playlist
        mask = (df_cases['Song'] == 'Grundaktivität') & (df_cases['Playlist'].astype(str).str.contains(playlist_key, case=False))
        
        if mask.any():
            subset = df_cases[mask]
            if agg_func == 'count': val = len(subset)
            else: val = subset[value_col].mean()
            
            # Schönen Namen generieren
            display_name = playlist_key.replace('audio', 'Audio ').replace('youtube', 'Youtube').capitalize()
            if 'shuffle' in playlist_key: display_name += " (Shuffle)"
            
            ga_stats.append({
                'Song': f"Grundaktivität ({display_name})",
                'Wert': val
            })
    
    # 3. Zusammenfügen
    if ga_stats:
        stats = pd.concat([stats, pd.DataFrame(ga_stats)], ignore_index=True)
        
    stats = stats.sort_values(by='Wert', ascending=False).reset_index(drop=True)
    return stats

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
        st.header("Daten Upload")
        uploaded_files = st.file_uploader("CSV Dateien", accept_multiple_files=True)
        cleaning_lvl = st.selectbox("Reinigungs-Level", CLEANING_LEVELS, index=0)
        
        st.markdown("### Cleaning Parameter")
        short_event_seconds = st.number_input(
            "Radius 'kurze' Events (s)", 0.5, 10.0, 2.5, 0.5,
            help="Zeitfenster um kurze Events."
        )
        reaction_buffer = st.number_input(
            "Reaktions-Puffer (s)", 0.0, 5.0, 1.0, 0.5,
            help="Erweitert den Ausschnitt rückwirkend (in die Vergangenheit)."
        )
        
        st.markdown("---")
        use_youtube = st.checkbox("Youtube-Daten anzeigen", value=False, 
                                  help="Wenn aktiviert, fließen Youtube Daten in die Statistiken ein.")

        if st.button("Daten einlesen", type="primary"):
            if not uploaded_files:
                st.error("Bitte Dateien hochladen.")
            else:
                participants = parse_uploaded_files(uploaded_files)
                process_data(participants, cleaning_lvl, short_event_seconds, reaction_buffer)

    if st.session_state['master_data'] is not None:
        df_full = st.session_state['master_data']
        
        if use_youtube:
            analysis_df = df_full.copy()
        else:
            # Alles entfernen, was "youtube" im Playlist-Namen hat (filtert 'youtube' und 'youtube_shuffle')
            analysis_df = df_full[~df_full['playlist'].str.contains('youtube', case=False)].copy()

        st.markdown("---")
        # -----------------------------------
        # ANALYSE PARAMETER
        # -----------------------------------
        st.header("Einstellungen")
        
        with st.container():
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Toleranz-Logik")
                # Auswahlmenü für die 3 Methoden
                calc_method = st.radio(
                    "Berechnungsmethode wählen:",
                    ["A) Absolut (Ohne Anpassung)", 
                     "B) Relativ (Dynamische Std-Abw.)", 
                     "C) Z-Score (Daten-Normalisierung)"],
                    index=2,
                    help="A: Fixer Wert für alle.\nB: Toleranz passt sich der Signalstärke an.\nC: Daten werden auf Standardnormalverteilung transformiert."
                )
                
                # Eingabefelder je nach Methode
                if "Absolut" in calc_method:
                    tol_val_input = st.number_input("Absoluter Schwellenwert", value=0.5, step=0.1)
                    st.caption("Verwendet Original-Daten & fixen Threshold.")
                    
                elif "Relativ" in calc_method:
                    std_factor = st.slider("Faktor der Standardabweichung (σ)", 0.1, 2.0, 0.5, 0.1)
                    st.caption(rf"Threshold = $\sigma \times {std_factor}$ (Daten bleiben original)")
                    
                else: # Z-Score
                    z_thresh = st.number_input("Z-Score Schwellenwert", value=0.5, step=0.1)
                    st.caption(r"Daten werden transformiert ($Z = \frac{x - \mu}{\sigma}$). Threshold gilt auf Z-Werten.")
                

            with col2:
                st.subheader("Impact Error Definition")
                high_impact_rank = st.number_input("Min. Rang-Diff für Inklusion in Impact Analyse", 
                                                 min_value=1, max_value=5, value=2)

        # -----------------------------------------------------------
        # ZENTRALE HILFSFUNKTION
        # -----------------------------------------------------------
        def apply_tolerance_logic(data_series):
            """
            Nimmt eine Serie von Messwerten und gibt zurück:
            1. Die (ggf. transformierten) Daten
            2. Den berechneten Schwellenwert für diese Daten
            """
            values = np.array(data_series)
            
            # Methode A: Absolut
            if "Absolut" in calc_method:
                return values, tol_val_input
            
            # Methode B: Relativ (Dynamischer Threshold auf Originaldaten) 
            elif "Relativ" in calc_method:
                if len(values) < 2: return values, 0.001
                sigma = np.std(values)
                return values, (sigma * std_factor)
            
            # Methode C: Z-Score (Daten transformieren)
            else:
                if len(values) < 2 or np.std(values) == 0: return values, 0.001 # Fallback
                z_values = (values - np.mean(values)) / np.std(values)
                return z_values, z_thresh

        st.markdown("---")

        tab_fein, tab_global, tab_meta, tab_verlauf = st.tabs([
            "Feingranulare Analyse", 
            "Globale Analyse", 
            "Meta-Daten Analyse", 
            "Verlaufsanalyse"
        ])

        # ==========================================================================================
        # TAB 1: FEINGRANULARE ANALYSE
        # ==========================================================================================
        with tab_fein:
            metric = st.selectbox(
                "Wähle die zu analysierende Metrik:", 
                METRICS, 
                index=1
            )

            st.header(f"Analyse: {metric}")

            with st.expander("📘 Methodik: Signalverarbeitung & Berechnung", expanded=False):
                st.markdown("### Der Weg vom Rohsignal zum Score")
                st.markdown(
                    "Die Berechnung erfolgt in einer dreistufigen Pipeline. Ziel ist es, aus dem komplexen Wellensignal "
                    "einen einzigen, robusten Zahlenwert (Score) pro Song zu extrahieren."
                )

                c_math1, c_math2 = st.columns(2)
                with c_math1:
                    st.markdown("**Schritt 1: Transformation (Zeit $\\to$ Frequenz)**")
                    st.markdown(
                        "Mittels der **Welch-Methode** wird das Signal in Fenster zerlegt und transformiert. "
                        "Wir berechnen die Fläche unter der Kurve (Integral) für ein Frequenzband (z.B. Alpha 8-13 Hz)."
                    )
                    st.latex(r"P_{Band} = \int_{f_{min}}^{f_{max}} PSD(f) \, df")

                with c_math2:
                    st.markdown("**Schritt 2: Räumliche Mittelung (Raum)**")
                    st.markdown(
                        "Das EEG liefert 8 verschiedene Werte (Kanäle). Um einen globalen Zustand zu erhalten, "
                        "bilden wir den Durchschnitt der Band-Power über alle Elektroden."
                    )
                    st.latex(r"P_{Global} = \frac{1}{N_{Ch}} \sum_{i=1}^{8} P_{Ch_i}")

                st.markdown("---")
                st.markdown("**Schritt 3: Verhältnisbildung (Ratio Score)**")
                
                c_rat1, c_rat2 = st.columns([2, 1])
                with c_rat1:
                    st.markdown(
                        """
                        Der finale Score ist kein absoluter Spannungswert, sondern ein **Verhältnis (Quotient)**.
                        Wir teilen die im Schritt 2 berechnete globale Power des einen Bandes durch die des anderen.
                        
                        *Beispiel für Alpha/Beta Ratio:*
                        """
                    )
                with c_rat2:
                    st.latex(r"Score = \frac{P_{Global}(\alpha)}{P_{Global}(\beta)}")

                st.markdown("---")
                st.markdown("#### Legende der Variablen")
                st.markdown(
                    """
                    | Symbol | Bedeutung | Erläuterung |
                    | :--- | :--- | :--- |
                    | $PSD(f)$ | Power Spectral Density | Die Stärke des Signals bei einer spezifischen Frequenz $f$. |
                    | $f_{min}, f_{max}$ | Frequenzgrenzen | Z.B. für Alpha: 8 Hz bis 13 Hz. |
                    | $P_{Band}$ | Band Power (Kanal) | Die gesamte Energie eines Frequenzbandes auf einem einzelnen Kanal. |
                    | $P_{Ch_i}$ | Power Kanal $i$ | Der $P_{Band}$-Wert für den $i$-ten EEG-Kanal. |
                    | $N_{Ch}$ | Anzahl Kanäle | Hier $N=8$ (EEG 1 bis EEG 8). |
                    | $P_{Global}$ | Globale Power | Der Durchschnittswert des Frequenzbandes über den ganzen Kopf. |
                    | $\\alpha, \\beta$ | Frequenzbänder | Alpha (Entspannung) und Beta (Fokus/Stress). |
                    """
                )

                st.markdown("---")
                st.markdown("## Methodische Begründung")
                
                st.markdown("##### Warum Welch-Methode? (Zeit-Dimension)")
                st.markdown(
                    """
                    Ein einfaches Periodogramm (FFT über den ganzen Song) wäre extrem "verrauscht" (hohe Varianz). 
                    Die Welch-Methode glättet das Spektrum statistisch, indem sie den Durchschnitt vieler kurzer Momente bildet. 
                    Das macht das Ergebnis robuster gegen kurzzeitige Ausreißer.
                    """
                )
                
                st.markdown("##### Warum Mittelwert über alle Kanäle? (Raum-Dimension)")
                st.markdown(
                    """
                    Da wir keine spezifische Lokalisierung (z.B. "nur Frontallappen") untersuchen, sondern die 
                    allgemeine Entspannungswirkung von Musik, nutzen wir den Durchschnitt aller 8 Elektroden.
                    * **Vorteil 1:** Reduziert den Einfluss lokaler Artefakte (z.B. wenn Elektrode 3 schlechten Kontakt hat).
                    * **Vorteil 2:** Erfasst eine systemweite Antwort des Gehirns statt nur lokaler Phänomene.
                    * **Vorteil 3:** Geringere Komplexität/Aufwand für Bachelor.
                    """
                )
                
                st.markdown("#### Warum Ratios statt absoluter Power?")
                st.markdown(
                    r"""
                    Die absolute Signalstärke ($\mu V$) ist individuell verschieden (Schädeldicke, Hautleitfähigkeit). 
                    Ratios (z.B. Alpha/Beta) sind relative Werte und daher besser zwischen Personen vergleichbar.
                    """
                )
            
            # -----------------------------------
            # TOLERANZ LOGIK WIRKUNG
            # -----------------------------------
            with st.expander("📘 Methodik: Normalisierung & Threshold", expanded=False):
                st.markdown("### Zielsetzung")
                st.markdown(
                    "Ein einfacher Rang-Fehler (z.B. Platz 1 statt 3) sagt nichts über die Schwere des Irrtums aus. "
                    "Wir wollen Fehler bestrafen, die **sowohl** im Ranking falsch sind **als auch** physiologisch stark abweichen."
                )
                
                st.markdown("---")
                st.markdown("### Die Formel (Weighted Score)")
                st.markdown("Der Score für einen Song berechnet sich aus zwei Komponenten:")
                st.latex(r"Score = \Delta Rank \times \Delta Physio")
                
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    st.markdown("**1. Der Rang-Fehler**")
                    st.latex(r"\Delta Rank = |Rank_{Subj} - Rank_{EEG}|")
                    st.caption("Bestraft die falsche Platzierung.")
                with col_f2:
                    st.markdown("**2. Die physiologische Distanz**")
                    st.latex(r"\Delta Physio = \begin{cases} |Val - Val_{Ref}|, & \text{wenn } > Threshold \\ 0, & \text{sonst} \end{cases}")
                    st.caption("Bestraft die Stärke der Signal-Abweichung.")
                
                st.markdown("---")
                st.markdown("### Die Rolle des Z-Score")
                st.markdown(
                    "Um den **Weighted Score** zwischen verschiedenen Probanden vergleichbar zu machen "
                    "(„Wer hat das schlechteste Körpergefühl?“), müssen die physiologischen Messwerte normalisiert werden. "
                    "Dazu muss der **Z-Score** verwendet werden."
                )
                
                st.latex(r"Z = \frac{x - \mu}{\sigma}")
                st.markdown(
                    r"""
                    * $x$: Der gemessene Wert (z.B. Alpha/Beta Ratio eines Songs).
                    * $\mu$: Der persönliche Durchschnittswert des Probanden in dieser Session.
                    * $\sigma$: Die persönliche Standardabweichung (Schwankungsbreite) des Probanden.
                    """
                )
                
                st.info(
                    "💡 **Effekt:** Ein Z-Score von **2.0** bedeutet bei jedem Probanden dasselbe: "
                    "„Dieser Song war 2 Standardabweichungen stärker als der Durchschnitt“ – egal ob der Proband "
                    "generell starke (50 µV) oder schwache (5 µV) Hirnströme hat."
                )

                st.markdown("---")
                st.markdown("### Entscheidungshilfe: Relativ vs. Z-Score")
                
                c_dec1, c_dec2 = st.columns(2)
                with c_dec1:
                    st.markdown("**B) Relativ (Dynamische Std-Abw.)**")
                    st.markdown("*Einsatz: Individuelle Betrachtung*")
                    st.markdown(
                        rf"Nutzt die Original-Werte (µV/Ratios). Der Threshold passt sich dem individuellen Probanten dynamisch an ($Threshold = \sigma \times Faktor$)."
                    )
                    st.markdown("✅ Intuitiv verständlich beim Vergleich der Analyse zwischen Songs eines Probanten innerhalb einer Session.")
                    st.markdown("❌ Nicht gut für Vergleiche zwischen Probanten.")

                with c_dec2:
                    st.markdown("**C) Z-Score (Normalisierung)**")
                    st.markdown("*Einsatz: Probanten-Vergleiche*")
                    st.markdown(
                        "Transformiert Daten auf eine Einheitsskala. Der Threshold ist fix (z.B. 0.5), da die Daten genormt sind."
                    )
                    st.markdown("✅ Macht Probanten Messwerte vergleichbar.")
                    st.markdown("❌ Abstraktere Werte.")

                st.markdown("---")
                st.markdown("### Warum bleiben Spearman & Kendall gleich bei Toleranz-Logik Wechsel zwischen B) und C)?")
                st.markdown(
                    r"Es mag überraschen, dass der Wechsel zwischen **Relativ** und **Z-Score** keinen Einfluss auf die "
                    "Korrelationswerte ($r_s$, $\\tau$) hat. Dies liegt an der mathematischen Natur von Rang-Analysen:"
                )
                st.markdown(
                    "**1. Ranking ist die Normalisierung:**\n"
                    r"Spearman und Kendall rechnen **nicht** mit absoluten Messwerten (µV oder Z-Scores), sondern ausschließlich "
                    "mit Platzierungen (Rängen). Sobald Daten in Ränge umgewandelt werden (Platz 1, 2, 3...), wird die absolute "
                    "Signalstärke ('Abweichung') effektiv entfernt. Ein riesiger Signal-Unterschied führt zum gleichen Rang-Abstand "
                    "wie ein kleiner Unterschied, solange die Reihenfolge gleich bleibt."
                )
                st.markdown(
                    "**2. Der Threshold skaliert mit:**\n"
                    r"Ob wir den Threshold an das Signal anpassen (Relativ: Threshold wächst mit $\sigma$) oder das Signal an den "
                    r"Threshold anpassen (Z-Score: Signal schrumpft durch $\sigma$) – das **Verhältnis** bleibt mathematisch identisch."
                )

                tol_cases = []
                if not analysis_df.empty:
                    for _, grp in analysis_df.groupby(['participant', 'playlist']):
                        # 1. Daten holen
                        raw_vals = grp[metric].values
                        
                        # 2. Dynamische Logik anwenden
                        vals, current_thresh = apply_tolerance_logic(raw_vals)
                        
                        sort_idx = np.argsort(vals)[::-1] # Absteigend
                        sorted_vals = vals[sort_idx]
                        sorted_names = grp['item'].iloc[sort_idx].values
                        
                        for k in range(len(sorted_vals)-1):
                            diff = abs(sorted_vals[k] - sorted_vals[k+1])
                            
                            if diff < current_thresh:
                                tol_cases.append({
                                    'Proband': grp['participant'].iloc[0], 
                                    'Playlist': grp['playlist'].iloc[0],
                                    'Items': f"{sorted_names[k]} ↔ {sorted_names[k+1]}", 
                                    'Diff': diff,
                                    'Threshold (Individuell)': current_thresh,
                                    'Ausnutzung (%)': (diff / current_thresh) * 100 if current_thresh > 0 else 0
                                })
                                
                    if tol_cases: 
                        st.dataframe(pd.DataFrame(tol_cases).style.format({'Diff': '{:.4f}', 'Threshold (Individuell)': '{:.4f}', 'Ausnutzung (%)': '{:.1f}%'}).map(color_tolerance_usage, subset=['Ausnutzung (%)']))
                    else: 
                        st.info(f"Keine Fälle gefunden, die den Toleranzbereich unterschreiten.")

            with st.expander("📘 Methodik: Rangzuordnung", expanded=False):
                st.markdown("### Zielsetzung")
                st.markdown(
                    "Klassische Korrelation bestraft jede Abweichung von Rängen unabhängig davon ob man überhaupt einen Subjektiven unterschied zwischen den Entspannungsgrad der Lieder merken kann.")
                st.markdown(
                    "**Wir wollen also prüfen:** Stimmt die **subjektive Reihenfolge** mit der **objektiven Reihenfolge** überein, wenn wir den Probanten eine reallistische Toleranz zugestehen. (Da wir nicht so feingranular unterscheiden können wie das EEG.)"
                )

                st.markdown("---")
                st.markdown("### Vergleich der Methoden")
                col_m1, col_m2, col_m3 = st.columns(3)
                with col_m1:
                    st.markdown("**Standard**")
                    st.caption("Mathematisch strikt")
                    st.write("Kein Toleranzbereich. User Ranking wird mit EEG Ranking **strikt verglichen.**")
                with col_m2:
                    st.markdown("**Fair (Tie)**")
                    st.caption("Konservativ realistisch")
                    st.write("Werte im Toleranzbereich erhalten den **gleichen Durchschnittsrang**.")
                    st.write("Jedoch ist es auch mit der Methode nicht möglich einen Wert von 1.0 zu erreichen, selbst wenn ein Unübereinstimmung 'verbessert' wird.")
                with col_m3:
                    st.markdown("**Sehr Fair (Opt)**")
                    st.caption("Optimistisch plausibel. Keine Globale Optimierung.")
                    st.write("Bei Gleichstand wird die Reihenfolge so gewählt, dass sie dem User-Ranking am nächsten kommt.")
                    st.write("Greedy Algorithmus (von oben nach unten), sodass der beste Wert noch als 'Anker' fungiert. Alle Plätze im Threshold-Cluster müssen zum Anker passen.")

                st.divider()

                st.subheader("Methodische Einordnung & Abgrenzung")

                c_val, c_over = st.columns(2)

                with c_val:
                    st.info("✅ **Warum 'Sehr Fair' wertvoll ist**\n(Best-Case Szenario / Plausibilität)")
                    st.markdown(
                        """
                        Diese Methode dient dazu, **Messrauschen und Kontext-Artefakte** (z.B. Schlechte Position der Elektroden, schlechte Verbindung, nicht erfasste Ablenkung während der Aufnahme) zu minimieren und ein **Optimalen Kontext** zu simulieren.
                        
                        **Der strategische Hintergrund (Hypothesen-Test):**
                        Wir wählen bewusst einen **optimistischen Ansatz** (Steel-Manning). Da eine Hauptthese der Arbeit ist, dass subjektive Einschätzungen alleine für die Erfassung von Effekten, ausgelöst durch auditiven Reize, zu ungenau sind, stärken wir die Beweiskraft: 
                        *Wenn die Korrelation selbst unter diesem 'Best-Case-Szenario' (Sehr Fair) niedrig bleibt, ist dies ein umso stärkerer Beweis für die Unzuverlässigkeit subjektiver Wahrnehmung.*

                        **Die Argumente:**
                        * **Kompensation von Unschärfe:** Da wir nicht exakt wissen, ab welcher µV-Differenz ein Mensch *wirklich* einen Unterschied spürt, nutzen wir bei physiologischer Uneindeutigkeit (im Threshold) die subjektive Meinung als Stütze.
                        * **Tie-Breaker Logik:** Innerhalb des Toleranzbereichs gilt das EEG als "neutral". Die subjektive Wahrnehmung darf hier die Reihenfolge entscheiden, ohne dass es als Fehler gewertet wird.
                        * **Ergebnis:** Ein hoher Wert bei "Sehr Fair" bedeutet: Die User-Meinung ist zumindest **physiologisch plausibel** (wird nicht durch harte Daten widerlegt).
                        """
                    )

                with c_over:
                    st.error("🚫 **Warum ein Fair-Hybrid-Ansatz ('Cherry-Picking') nicht verwendet wird**")
                    st.markdown(
                        """
                        Man könnte versucht sein, eine Methode zu bauen, die **Standard** nutzt, wenn es passt (Lucky Hits behalten), und **Fair** nutzt, wenn es nicht passt (Fehler ignorieren).
                        
                        * **Das Problem (Data Dredging):** Man würde die Auswertungs-Regeln dynamisch ändern, um das Ergebnis künstlich zu verbessern.
                        * **Warum das falsch ist:** Wenn man Messdaten nur dann als "präzise" akzeptiert, wenn sie die eigene Meinung bestätigen (Standard), sie aber als "Rauschen" abtut, sobald sie widersprechen (Fair), testet man nicht mehr die Realität. Man betreibt **Confirmation Bias** (Bestätigungsfehler).
                        * **Unterschied zu 'Sehr Fair':** * *Sehr Fair* folgt einer **festen Regel** (Cluster-Optimierung basierend auf einem fixen Threshold).
                            * *Hybrid* folgt **opportunistischen Regeln** (Regelwerk ändert sich je nach gewünschtem Output).
                        * **Fazit:** Ein solcher Ansatz ist wissenschaftlich unsauber (Overfitting). Wir bleiben daher bei den konsistenten Modellen.
                        """
                    )
            # -----------------------------------
            # RANGKORRELATION
            # -----------------------------------
            st.subheader("Rangkorrelation")
            

            with st.expander("📘 Methodik: Spearman vs. Kendall", expanded=False):
                st.markdown(
                    "Beide Metriken messen die Ähnlichkeit von zwei Ranglisten (z.B. *Subjektive Meinung* vs. *EEG-Messung*). "
                    "Sie unterscheiden sich jedoch darin, was sie Aussagen und als 'Fehler' betrachten."
                )
                
                tab_spear, tab_kend, tab_eval = st.tabs(["Spearman (Distanz)", "Kendall (Paare)", "Bewertung (Output)"])
                
                # --- TAB 1: SPEARMAN ---
                with tab_spear:
                    st.markdown("### Spearman ($r_s$): Der Abstands-Messer")
                    st.info(
                        "**Philosophie:** \"Wie **weit** liegen die Platzierungen auseinander?\" "
                        "Ein Fehler um 5 Plätze wird viel härter bestraft als ein Fehler um 1 Platz."
                    )
                    
                    st.markdown("**1. Die Formel**")
                    st.markdown("Spearman basiert auf der **quadrierten Differenz** ($d^2$) der Ränge.")
                    st.latex(r"r_s = 1 - \frac{6 \sum d_i^2}{n(n^2 - 1)}")
                    st.caption(r"$d_i$: Differenz zwischen Rang A und Rang B für Song i | $n$: Anzahl der Songs")
                    
                    st.markdown("**2. Vorgehensweise**")
                    st.markdown(
                        """
                        1.  Wandle die Messwerte (µV) in ein Verhältnis um. Anhand der Verhältnisse erstelle ein Ranking (1. [höchste Zahl] bis n. [kleinste Zahl]).
                        2.  Bilde für jeden Song die Differenz: $d = Rang_{User} - Rang_{EEG}$.
                        3.  **Quadriere** diese Differenz ($d^2$). *Das sorgt dafür, dass große Fehler (High Impact) überproportional stark ins Gewicht fallen.*
                        4.  Summiere alle $d^2$ auf und setze sie in die Formel ein.
                        """
                    )

                # --- TAB 2: KENDALL ---
                    with tab_kend:
                        st.markdown("### Kendall ($\\tau$): Die Paar-Logik")
                        st.info(
                            "**Philosophie:** Anstatt auf die genauen Platznummern zu schauen, zerlegt Kendall die gesamte Playlist "
                            "in alle möglichen Zweier-Pärchen und prüft jedes einzeln auf Widersprüche."
                        )
                        st.markdown("#### Berechnung")
                        c_def1, c_def2 = st.columns(2)
                        with c_def1:
                            st.success("**Konkordantes Paar (+1)**")
                            st.caption("Einigkeit")
                            st.markdown("User und EEG sortieren das Paar gleich.")
                            st.markdown("*Bsp: User sagt A > B. EEG sagt auch A > B.*")
                        with c_def2:
                            st.error("**Diskordantes Paar (-1)**")
                            st.caption("Widerspruch (Inversion)")
                            st.markdown("Die Reihenfolge ist vertauscht.")
                            st.markdown("*Bsp: User sagt A > B. EEG sagt aber B > A.*")

                        st.markdown("**Formel:**")
                        st.latex(r"\tau = \frac{Konkordant - Diskordant}{\text{Alle Paare}}")

                        st.latex(r"\text{Alle Paare} = \frac{n(n-1)}{2}")
                                
                        st.markdown("**Die Wahrscheinlichkeits-Aussage (Der 'Wett-Faktor')**")
                        st.markdown(
                            "Aus dem $\\tau$-Wert lässt sich direkt berechnen, wie oft User und EEG einer Meinung sind, "
                            "wenn man zwei beliebige Songs vergleicht:"
                        )
                        st.latex(r"P(\text{Übereinstimmung}) = \frac{\tau + 1}{2}")
                        
                        st.warning("Bei der Kendall-Wahrscheinlichkeit bedeutet 'Übereinstimmung': Der Song wurde vom User genauso wie vom EEG als entspannter bewertet (platziert) als der andere.")

                    # --- TAB 3: INTERPRETATION ---
                        with tab_eval:
                            st.markdown("### Welche Frage beantwortet welcher Algorithmus?")
                            st.markdown(
                                "Bevor man auf die reinen Zahlen schaut, muss man wissen, was der jeweilige Algorithmus über die "
                                "Fähigkeiten des Probanden aussagt:"
                            )
                            
                            # Oben: Die Aussagekraft (Spearman vs. Kendall)
                            c_algo1, c_algo2 = st.columns(2)
                            with c_algo1:
                                st.info("**1. Spearman ($r_s$): Globale Hierarchie**")
                                st.markdown("*Fokus: Distanz & Schwere Fehler*")
                                st.markdown(
                                    "Beantwortet die Frage:\n"
                                    "**\"Wie gut kann der Proband mehrere Lieder insgesamt voneinander unterscheiden?\"**\n\n"
                                    "Ein hoher positiver Wert bedeutet, dass der Proband ein Gefühl für das **Gesamtbild** hat. Also in unserem Fall die Fähigkeit besitzt gut den Entspannungsgrad mehrere Songs (**mit BL**) untereinander zu unterscheiden."
                                )
                                st.markdown("**→** Ranking des Probands hat eine hohe Übereinstimmung (positive Korrelation) mit dem des EEGs.")
                                
                            with c_algo2:
                                st.info("**2. Kendall ($\\tau$): Lokale Unterscheidung**")
                                st.markdown("*Fokus: Konsistenz & Wahrscheinlichkeit*")
                                st.markdown(
                                    "Beantwortet die Frage:\n"
                                    "**\"Wie gut kann der Proband den Effekt zwischen zwei Liedern unterscheiden?\"**\n\n"
                                    "Ein hoher Wert bedeutet, dass der Proband im **direkten Vergleich** (A vs. B) zuverlässig erkennt, "
                                    "welcher Song physiologisch entspannender ist."
                                )
                                st.markdown("**→** Der Probands hat eine hohe Übereinstimmung (positive Korrelation) mit dem dem EEG, wenn es darum geht den Entspannungsgrad zwischen zwei Songs zu unterscheiden.")

                            st.markdown("### 📊 Bewertung der Korrelations-Werte")
                            st.markdown(
                                "Die Interpretation der Werte erfolgt vor dem Hintergrund, dass hier **subjektive Rankings (1-6)** "
                                "mit **physiologischen Daten (EEG)** verglichen werden. Da dies oft mit 'Rauschen' behaftet ist, "
                                "sind bereits moderate Korrelationen als substanziell zu bewerten."
                            )

                            col_pos, col_neg = st.columns(2)

                            # --- Linke Spalte: Positiv (Treffer) ---
                            with col_pos:
                                st.success("🟢 **Positiver Bereich**\n(Korrekte Wahrnehmung)")
                                st.caption("Je höher der Wert, desto besser passt das Gefühl zu den Daten.")
                                
                                st.markdown("**1. Spearman ($r_s$)**")
                                st.markdown(
                                    """
                                    * **> +0.50:** Sehr starker Zusammenhang (Exzellent).
                                    * **+0.30 bis +0.50:** Starker Effekt (nach neuerer Empirie) / Moderat (klassisch).
                                    * **+0.10 bis +0.30:** Kleiner bis moderater Effekt.
                                    * **< +0.10:** Kein nennenswerter Zusammenhang.
                                    """
                                )
                                
                                st.divider()
                                
                                st.markdown("**2. Kendall ($\\tau$)**")
                                st.markdown(
                                    "Da $\\tau$ mathematisch bedingt meist kleiner ist als $r_s$, "
                                    "sind die Schwellenwerte hier **ca. 33% niedriger** angesetzt (Greiner's Relation)."
                                )
                                st.markdown(
                                    """
                                    * **> +0.34:** Sehr starker Effekt.
                                    * **+0.20 bis +0.34:** Stark / Moderat (Substanzieller Zusammenhang).
                                    * **+0.07 bis +0.20:** Klein / Schwach.
                                    """
                                )

                            # --- Rechte Spalte: Negativ (Fehler) ---
                            with col_neg:
                                st.error("🔴 **Negativer Bereich**\n(Invertierte Wahrnehmung)")
                                st.caption("Das subjektive Gefühl widerspricht den objektiven Messdaten.")
                                
                                st.markdown("**1. Spearman ($r_s$)**")
                                st.markdown(
                                    """
                                    * **-0.10 bis -0.30:** Leichte Abweichung / Rauschen.
                                    * **-0.30 bis -0.50:** Deutliche Fehlattribution.
                                    * **< -0.50:** Starke Diskrepanz (Gegenteilige Wahrnehmung).
                                    """
                                )
                                
                                st.divider()
                                
                                st.markdown("**2. Kendall ($\\tau$)**")
                                st.markdown(
                                    """
                                    * **> -0.20:** Rauschen / Zufall.
                                    * **-0.20 bis -0.34:** Deutliche Abweichung.
                                    * **< -0.34:** Starke Diskrepanz.
                                    """
                                )

                            # --- Wissenschaftliche Quellen & Methodik ---
                            st.markdown("---")
                            st.markdown("#### Wissenschaftliche Fundierung der Schwellenwerte")

                            with st.expander("Details zu Quellen und Methodik anzeigen", expanded=False):
                                st.markdown(
                                    """
                                    **1. Interpretation von Spearman ($r_s$):**
                                    * Traditionell gelten nach Cohen (1988) Werte von **0.1, 0.3 und 0.5** als klein, mittel und groß.
                                    * Neuere Meta-Analysen in der Differentiellen Psychologie zeigen jedoch, dass diese Hürden oft unrealistisch hoch sind. Gignac & Szodorai (2016) schlagen vor, bereits Werte ab **0.20 als typisch** und ab **0.30 als groß** zu bewerten.
                                    * *Im Kontext dieser Arbeit (Subjektiv vs. Physiologisch) wird daher der Bereich 0.30–0.50 bereits als substanzielles Ergebnis gewertet.*

                                    **2. Umrechnung zu Kendall ($\\tau$):**
                                    * Kendalls Tau nimmt bei gleicher Assoziationsstärke systematisch niedrigere Werte an als Spearman.
                                    * Unter der Annahme einer bivariaten Normalverteilung gilt die **Greiner-Relation**: $r_s \\approx \\frac{3}{2}\\tau$ bzw. $\\tau \\approx \\frac{2}{3}r_s$.
                                    * Daraus ergeben sich die angepassten Schwellenwerte (z.B. $r_s=0.50 \\rightarrow \\tau \\approx 0.34$).
                                    
                                    **Quellenverzeichnis:**
                                    * *Cohen, J. (1988).* Statistical Power Analysis for the Behavioral Sciences. Erlbaum.
                                    * *Gignac, G. E., & Szodorai, E. T. (2016).* Effect size guidelines for individual differences researchers. *Personality and Individual Differences, 102*, 74–78.
                                    * *Gilpin, A. R. (1993).* Table for Conversion of Kendall's Tau to Spearman's Rho. *Educational and Psychological Measurement, 53(1)*, 87-92.
                                    * *Fredricks, G. A., & Nelsen, R. B. (2007).* On the relationship between Spearman's rho and Kendall's tau. *Journal of Statistical Planning and Inference, 137(7)*, 2143-2150.
                                    """
                                )

            with st.expander("❓ Warum sind 'faire' Werte oft schlechter als 'Standard'?", expanded=False):
                st.markdown("### Das Paradoxon")
                st.markdown(
                    "Es wirkt intuitiv falsch: Wenn man dem EEG eine **Toleranz** zugesteht (Modus 'Fair'), "
                    "sollten die Ergebnisse doch eigentlich besser werden, oder? "
                    "Dass sie oft **sinken**, liegt an der Entfernung von **Zufallstreffern (Lucky Hits)**."
                )
                st.markdown("Diese Zufallstreffer resultieren aus der Diskrepanz zwischen maschineller Messgenauigkeit und menschlicher Wahrnehmung: Während Maschinen feinste Nuancen registrieren, können Menschen subjektiv nur gröbere Unterschiede im Entspannungsgrad feststellen.")

                st.divider()

                st.markdown("### Der Grund für den 'Punktabzug'")
                st.markdown(
                    "Wir gehen von folgendem Szenario aus: **Der User hat eine strikte Meinung** (Song A ist Platz 1, Song B ist Platz 2). "
                    "Er vergibt also *keine* Unentschieden."
                )

                tab_why_k, tab_why_s, tab_sehr_fair = st.tabs(["Bei Kendall (Punkte)", "Bei Spearman (Distanz)", "Bei Sehr Fair (Opt)"])

                with tab_why_k:
                    st.info("**Kendall ($\\tau$): Das Prinzip der 'Wette'**")
                    st.markdown(
                        """
                        Kendall schaut sich Paare an und vergibt Punkte:
                        * **+1 Punkt:** Übereinstimmung (Konkordant)
                        * **-1 Punkt:** Widerspruch (Diskordant)
                        * **0 Punkte:** Unentschieden (Tie)
                        """
                    )
                    
                    c_k1, c_k2 = st.columns(2)
                    with c_k1:
                        st.write("**Standard (Lucky Hit)**")
                        st.caption("Rauschen zufällig richtig")
                        st.markdown("- User: A > B")
                        st.markdown("- EEG: A > B (z.B. um 0.01 µV)")
                        st.markdown("→ **Ergebnis: +1 Punkt**")
                    
                    with c_k2:
                        st.write("**Fair (Tied Ranks)**")
                        st.caption("Rauschen ignoriert")
                        st.markdown("- User: A > B")
                        st.markdown("- EEG: A = B (Tie)")
                        st.markdown("→ **Ergebnis: 0 Punkte**")
                        
                    st.error(
                        "**Fazit:** Durch 'Fair' verlierst du den +1 Punkt aus dem Zufallstreffer. "
                        "Du bekommst zwar keinen Minuspunkt, aber da der Zähler in der Formel kleiner wird, sinkt die Korrelation."
                    )

                with tab_why_s:
                    st.info("**Spearman ($r_s$): Das Prinzip der 'Straf-Distanz'**")
                    st.markdown(
                        r"""
                        Spearman startet bei 1.0 (Perfekt) und zieht für jeden Fehler etwas ab. 
                        Der Abzug basiert auf der **quadrierten Differenz** ($d^2$) zwischen den Rängen.
                        
                        $$Abzug \propto (Rank_{User} - Rank_{EEG})^2$$
                        """
                    )

                    c_s1, c_s2 = st.columns(2)
                    with c_s1:
                        st.write("**Standard (Lucky Hit)**")
                        st.markdown("- User Rank: **1**")
                        st.markdown("- EEG Rank: **1**")
                        st.latex(r"d = |1 - 1| = 0")
                        st.markdown("→ **Abzug: 0**")
                        st.caption("(Perfekt)")
                    
                    with c_s2:
                        st.write("**Fair (Tied Ranks)**")
                        st.markdown("- User Rank: **1**")
                        st.markdown("- EEG Rank: **1.5** (Mittelwert aus 1+2)")
                        st.latex(r"d = |1 - 1.5| = 0.5")
                        st.latex(r"d^2 = 0.25")
                        st.markdown("→ **Abzug: > 0**")
                        st.caption("(Fehler steigt)")

                    st.error(
                        "**Fazit:** Da der User '1' sagt und das EEG '1.5' (Unentschieden), entsteht mathematisch eine Distanz von 0.5. "
                        "Diese Distanz verringert die Korrelation im Vergleich zum perfekten Zufallstreffer (Distanz 0)."
                    )

                with tab_sehr_fair:
                    st.markdown("### Kann 'Sehr Fair' (Optimiert) schlechter sein als 'Standard'?")
                    st.markdown("**Praktisch: Extrem selten | Theoretisch: Ja.**")
                    
                    st.markdown(
                        "Der Modus 'Sehr Fair' versucht, die Ränge innerhalb des Toleranzbereichs so zu sortieren, "
                        "dass sie dem User entsprechen (Distanz = 0). Dennoch kann es in **mathematischen Randfällen** zu einer Verschlechterung kommen."
                    )

                    st.warning("⚠️ Der Grund: 'Greedy' Cluster-Bildung")
                    st.markdown(
                        """
                        **Das Szenario:**
                        Wenn der Threshold ungünstig gewählt ist, können Songs in ein Cluster "gesaugt" werden, die global betrachtet nicht zusammengehören (Kettenreaktion: A nah an B, B nah an C $\\to$ A, B, C in einem Cluster).
                        
                        * **Standard:** Hatte zufällig durch Rauschen noch eine gewisse Distanz zwischen A und C gewahrt (Lucky Hit in der globalen Struktur).
                        * **Sehr Fair:** Zwingt alle in einen Topf und vergibt neue Ränge (1, 2, 3) basierend auf lokaler Optimierung. Dabei kann in seltenen Fällen eine zufällig korrekte *globale* Distanz-Information der Standard-Methode verloren gehen.
                        """
                    )

                    st.info(
                        "**Bis zu welchem Grad? (Magnitude)**\n\n"
                        "Die Verschlechterung ist fast immer **marginal** und bewegt sich meist im Bereich der **3. oder 4. Nachkommastelle** (z.B. 0.452 $\\to$ 0.449).\n"
                        "Es handelt sich dabei nicht um einen methodischen Fehler, sondern um ein statistisches Artefakt der Cluster-Bildung."
                    )

            res_spearman = []
            res_kendall = []
            
            if not analysis_df.empty:
                for pid in analysis_df['participant'].unique():
                    p_data = analysis_df[analysis_df['participant'] == pid]
                    
                    # Dynamische Logik
                    raw_vals = p_data[metric].values
                    vals, current_thresh = apply_tolerance_logic(raw_vals)
                    temp_df = p_data.copy()
                    temp_df[metric] = vals
                    
                    # Berechnung BEIDER Varianten
                    (s_std, s_tie, s_opt), (k_std, k_tie, k_opt) = calculate_correlation_variants(temp_df, metric, 'subj_rank', current_thresh)
                    
                    res_spearman.append({
                        'Proband': pid, 'Standard': s_std, 'Fair (Tie)': s_tie, 'Sehr Fair (Opt)': s_opt
                    })
                    res_kendall.append({
                        'Proband': pid, 'Standard': k_std, 'Fair (Tie)': k_tie, 'Sehr Fair (Opt)': k_opt
                    })
                
                # TABS FÜR DUPLIZIERTE ANSICHT
                tab_s, tab_k = st.tabs(["Spearman (Rangabstand)", "Kendall (Paar-Wahrscheinlichkeit)"])
                
                with tab_s:
                    st.dataframe(
                        pd.DataFrame(res_spearman)
                        .set_index('Proband')
                        .style.format("{:.3f}", subset=['Standard', 'Fair (Tie)', 'Sehr Fair (Opt)'])
                        .background_gradient(cmap='RdYlGn', subset=['Standard', 'Fair (Tie)', 'Sehr Fair (Opt)'], vmin=-1, vmax=1))
                
                with tab_k:
                    # 1. Die normale Korrelations-Tabelle
                    st.markdown("###### Korrelations-Koeffizient (Tau)")
                    st.dataframe(
                        pd.DataFrame(res_kendall)
                        .set_index('Proband')
                        .style.format("{:.3f}", subset=['Standard', 'Fair (Tie)', 'Sehr Fair (Opt)'])
                        .background_gradient(cmap='RdYlGn', subset=['Standard', 'Fair (Tie)', 'Sehr Fair (Opt)'], vmin=-1, vmax=1)
                    )

                    st.markdown("---")
                    
                    # 2. Die neue Wahrscheinlichkeits-Tabelle
                    st.markdown("###### P(Übereinstimmung): Wahrscheinlichkeit der korrekten Identifizierung des höheren Entspannungsgrades zwischen zwei Songs.")

                    # Daten kopieren und umrechnen: P = (Tau + 1) / 2
                    df_prob = pd.DataFrame(res_kendall).set_index('Proband').copy()
                    cols = ['Standard', 'Fair (Tie)', 'Sehr Fair (Opt)']
                    df_prob[cols] = (df_prob[cols] + 1) / 2

                    st.dataframe(
                        df_prob.style
                        .format("{:.1%}", subset=cols) # Als Prozent formatieren (z.B. 75.0%)
                        .background_gradient(cmap='Blues', subset=cols, vmin=0.5, vmax=1.0) # Blau: Je dunkler, desto sicherer
                    )

            else: st.warning("Keine Daten.")

            # -----------------------------------
            # WEIGHTED HIGH IMPACT
            # -----------------------------------
            st.subheader("Physiologisch Impact Analyse")
            
            with st.expander("📘 Methodik: Berechnung des 'Weighted Impact Score'", expanded=False):
                st.markdown("### Das Prinzip: Nicht jeder Fehler ist gleich schwer")
                st.markdown(
                    "Ein einfacher Rang-Fehler (z.B. Platz 2 statt 1) sagt wenig aus. "
                    "Wir wollen wissen: **Hat sich der Proband nur leicht vertan oder lag er komplett daneben?**"
                )

                st.markdown("---")
                
                c_step1, c_step2 = st.columns(2)
                
                with c_step1:
                    st.markdown(r"**Schritt 1: Der Rang-Fehler ($\Delta Rank$)**")
                    st.markdown("Wie viele Plätze liegt die Einschätzung daneben?")
                    st.latex(r"\Delta Rank = |Rank_{Gefühl} - Rank_{EEG}|")
                    st.caption(r"Beispiel: Gefühl=1, EEG=3 $\Rightarrow$ Differenz = 2.")

                with c_step2:
                    st.markdown(r"**Schritt 2: Der physiologische Fehler ($\Delta Physio$)**")
                    st.markdown("Wie stark unterscheidet sich das Signal tatsächlich?")
                    st.latex(r"\Delta Physio = |Value_{Messung} - Value_{Referenz}|")
                    st.caption("Differenz zwischen dem gemessenen Wert und dem Wert, der auf dem erwarteten Rangplatz lag.")

                st.markdown("---")
                st.markdown("**Schritt 3: Der Gatekeeper (Toleranz)**")
                st.markdown(
                    r"Bevor wir Punkte vergeben, prüfen wir: **Ist der physiologische Unterschied überhaupt relevant?** "
                    r"Wenn die Differenz ($\Delta Physio$) kleiner ist als dein gewählter Schwellenwert (Threshold), "
                    "gilt der Fehler als 'physiologisch nicht wahrnehmbar' und der Score wird auf 0 gesetzt."
                )
                
                st.markdown("---")
                st.markdown("**Schritt 4: Die finale Formel**")
                st.markdown(
                    "Wenn die Toleranz überschritten wird, berechnet sich der Score als Produkt aus Rang-Irrtum und Signal-Stärke."
                )
                
                c_form1, c_form2 = st.columns([2, 1])
                with c_form1:
                    st.info(
                        r"**Logik:** Ein grober Rangfehler (z.B. 4 Plätze) bei einem sehr eindeutigen Signal (hohes $\Delta Physio$) "
                        "ergibt einen **extrem hohen Score** (High Impact Error)."
                    )
                with c_form2:
                    st.latex(r"Score = \Delta Rank \times \Delta Physio")

                st.markdown("---")
                st.markdown("### ⚠️ Besonderheit bei 'Fair' / 'Sehr Fair'")
                st.markdown("**Der doppelte Threshold-Effekt**")
                st.markdown(
                    r"""
                    In den Modi *Fair* und *Sehr Fair* wirkt der gewählte Toleranz-Wert (Threshold) gleich **zweifach** dämpfend auf das Ergebnis:
                    
                    1.  **Einfluss auf die Ränge (Vorstufe):** Der Threshold sorgt bereits bei der Erstellung der Ränge dafür, dass physiologisch ähnliche Songs denselben Rang erhalten (Tied Ranks). 
                        Dadurch wird $\Delta Rank$ oft bereits auf 0 reduziert.
                    2.  **Einfluss auf den Score (Gatekeeper):** Selbst wenn noch eine Rang-Differenz besteht, prüft der Threshold hier in Schritt 3 erneut, ob die Signal-Differenz signifikant genug für eine Bestrafung ist.
                    
                    **Fazit:** Ein Fehler wird in diesen Modi nur gezählt, wenn er sich **weder** durch die Rang-Toleranz **noch** durch die Signal-Toleranz erklären lässt. Dies filtert Rauschen extrem aggressiv.
                    """
                )

            col_mode, _ = st.columns([1, 2])
            with col_mode:
                impact_mode = st.selectbox(
                    "Berechnungs-Modus wählen:",
                    ["Standard", "Fair", "Sehr Fair"],
                    index=1, # Default auf Fair
                    key=f"impact_mode_selector_fein" 
                )

            weighted_cases = []
            if not analysis_df.empty:
                for (pid, playlist), group in analysis_df.groupby(['participant', 'playlist']):
                    # 1. Daten holen
                    raw_vals = group[metric].values
                    subj_ranks = group['subj_rank'].values
                    
                    # 2. Dynamische Logik (Toleranz bestimmen)
                    vals, current_thresh = apply_tolerance_logic(raw_vals)
                    
                    # 3. RANG BERECHNUNG JE NACH MODUS
                    if "Standard" in impact_mode:
                        calc_ranks = calculate_physio_ranks(vals, 0.0)
                        
                    elif "Sehr Fair" in impact_mode:
                        calc_ranks = np.zeros(len(vals))
                        sort_idx = np.argsort(vals)[::-1] 
                        sorted_values = vals[sort_idx]
                        
                        current_rank_counter = 1
                        i_idx = 0
                        while i_idx < len(vals):
                            cluster_indices = [sort_idx[i_idx]]
                            j = i_idx + 1
                            while j < len(vals) and abs(sorted_values[i_idx] - sorted_values[j]) < current_thresh:
                                cluster_indices.append(sort_idx[j])
                                j += 1
                                
                            available_ranks = list(range(current_rank_counter, current_rank_counter + len(cluster_indices)))
                            cluster_subj_ranks = subj_ranks[cluster_indices]
                            
                            sorted_rel_indices = np.argsort(cluster_subj_ranks)
                            
                            for k_idx, rel_idx in enumerate(sorted_rel_indices):
                                calc_ranks[cluster_indices[rel_idx]] = available_ranks[k_idx]
                                
                            current_rank_counter += len(cluster_indices)
                            i_idx = j
                        
                    else: # Fair (Default)
                        calc_ranks = calculate_physio_ranks(vals, current_thresh)

                    # Referenzwerte für Delta-Berechnung (Sortiert)
                    sorted_vals_desc = np.sort(vals)[::-1]
                    
                    temp_grp = group.copy().reset_index(drop=True)
                    temp_grp['transformed_val'] = vals
                    temp_grp['calc_rank'] = calc_ranks
                    
                    for idx, row in temp_grp.iterrows():
                        s_rank = int(row['subj_rank'])
                        c_rank = row['calc_rank']
                        val_act = row['transformed_val']
                        
                        ref_idx = min(s_rank - 1, len(sorted_vals_desc) - 1)
                        val_ref = sorted_vals_desc[ref_idx]
                        
                        rank_diff = abs(s_rank - c_rank)
                        val_diff = abs(val_act - val_ref)
                        
                        is_physio_error = val_diff >= current_thresh
                        
                        weighted_score = (rank_diff * val_diff) if is_physio_error else 0.0
                        
                        weighted_cases.append({
                            'Proband': row['participant'], 
                            'Playlist': row['playlist'],
                            'Song': row['item'], 
                            'Rang-Diff': round(rank_diff, 2),
                            'RankDiff_Int': int(round(rank_diff)),    
                            'Δ Messwert': round(val_diff, 4),
                            'Weighted Score': round(weighted_score, 4)
                        })

            if weighted_cases:
                w_df = pd.DataFrame(weighted_cases)
                
                hi_df = w_df[
                    (w_df['Weighted Score'] > 0) & 
                    (w_df['RankDiff_Int'] >= high_impact_rank)
                ].copy()

                total_users_analyzed = w_df['Proband'].nunique()
                total_playlists_analyzed = w_df.groupby(['Proband', 'Playlist']).ngroups
                
                total_hi_errors = len(hi_df)
                avg_hi_per_playlist = total_hi_errors / total_playlists_analyzed if total_playlists_analyzed > 0 else 0
                
                st.markdown("#### 📊 Analyse-Ergebnisse")
                
                kpi1, kpi2, kpi3, kpi4 = st.columns(4)
                kpi1.metric("Analysierte Probanden", total_users_analyzed)
                kpi2.metric("Analysierte Playlists", total_playlists_analyzed)
                kpi3.metric(f"Ø High-Impact Fehler", f"{avg_hi_per_playlist:.2f}", help="Pro Playlist")
                kpi4.metric("Total High-Impact Fehler", total_hi_errors, help=f"Anzahl aller Fehler mit Rang-Diff ≥ {high_impact_rank} und Score > 0.")
                
                st.markdown("---")
                
                c_chart, c_table = st.columns([1, 1])
                
                with c_chart:
                    st.markdown(f"**Verteilung der Fehler-Schwere (Ab Diff ≥ {high_impact_rank})**")
                    if not hi_df.empty:
                        dist_counts = hi_df['RankDiff_Int'].value_counts().reset_index()
                        dist_counts.columns = ['Rang-Differenz', 'Anzahl']
                        dist_counts = dist_counts[dist_counts['Rang-Differenz'] <= 5]

                        chart_dist = alt.Chart(dist_counts).mark_bar().encode(
                            x=alt.X('Rang-Differenz:O', title='Größe des Irrtums (Plätze)'),
                            y=alt.Y('Anzahl:Q', title='Häufigkeit'),
                            color=alt.Color('Rang-Differenz:O', legend=None, scale=alt.Scale(scheme='orangered')),
                            tooltip=['Rang-Differenz', 'Anzahl']
                        ).properties(height=250)
                        
                        st.altair_chart(chart_dist, width='stretch')
                    else:
                        st.info("Keine High-Impact Fehler unter den gewählten Kriterien.")

                with c_table:
                    st.markdown("**Signifikante Einzelfehler (Details):**")
                    relevant_errors = hi_df.sort_values(by='Weighted Score', ascending=False)
                    
                    if not relevant_errors.empty:
                        st.dataframe(
                            relevant_errors[['Proband', 'Playlist', 'Song', 'Rang-Diff', 'Weighted Score']]
                            .style.background_gradient(subset=['Weighted Score'], cmap='Reds'),
                            height=250
                        )
                    else:
                        st.info("Alle Rang-Fehler liegen innerhalb der physiologischen Toleranz.")

                st.markdown("---")
                st.info(
                "💡 **Wichtiger Hinweis zur Vergleichbarkeit:** "
                "Für einen Vergleich über mehrere Probanden hinweg {Gesamtübersicht} oder zwischen Probanten {Signifikaten Einzelfeher (Details)} sollte in den Einstellungen "
                "oben zwingend der **Z-Score** gewählt werden.\n\n"
                )
                
                st.markdown("**Gesamtübersicht (Ø Weighted Score aller Songs):**")
                stats_weighted = generate_detailed_stats(w_df, 'Weighted Score', agg_func='mean')
                st.dataframe(stats_weighted.style.background_gradient(subset=['Wert'], cmap='Reds').format({'Wert': '{:.4f}'}))

            else: st.success("Keine Daten für Berechnung.")

        # ==========================================================================================
        # TAB 2: GLOBALE ANALYSE
        # ==========================================================================================
        with tab_global:
            # -----------------------------------
            # GLOBALE META-ANALYSE (STUFEN-VERGLEICH)
            # -----------------------------------
            st.header("Globale Aussagen (Meta-Analyse)")
            
            st.info(
                "Hier aggregieren wir die Ergebnisse über alle Probanden hinweg. "
                "Wir vergleichen direkt, wie sich die Korrelation verbessert, wenn wir von der mathematisch strengen (Standard) "
                "zur physiologisch toleranten (Fair / Sehr Fair) Bewertung übergehen."
            )

            col_gs_set, _ = st.columns([1, 2])
            with col_gs_set:
                global_ws_mode = st.selectbox(
                    "Berechnungs-Modus für Fehler-Score (Item-Analyse):",
                    ["Standard", "Fair", "Sehr Fair"],
                    index=1, # Default auf Fair
                    help="Beeinflusst die Berechnung der 'Globalen Einschätzungs-Schwierigkeit' weiter unten."
                )

            meta_corr_s = []
            meta_corr_k = []
            meta_weighted_scores = []
            
            for m in METRICS:
                if not analysis_df.empty:
                    # A) KORRELATIONEN
                    s_std_l, s_fair_l, s_sfair_l = [], [], []
                    k_std_l, k_fair_l, k_sfair_l = [], [], []
                    
                    for pid in analysis_df['participant'].unique():
                        p_data = analysis_df[analysis_df['participant'] == pid]
                        
                        raw_vals = p_data[m].values
                        vals, current_thresh = apply_tolerance_logic(raw_vals)
                        temp_df = p_data.copy()
                        temp_df[m] = vals
                        
                        (s_std, s_tie, s_opt), (k_std, k_tie, k_opt) = calculate_correlation_variants(temp_df, m, 'subj_rank', current_thresh)
                        
                        if not np.isnan(s_std):
                            s_std_l.append(s_std); s_fair_l.append(s_tie); s_sfair_l.append(s_opt)
                            k_std_l.append(k_std); k_fair_l.append(k_tie); k_sfair_l.append(k_opt)
                    
                    if s_std_l:
                        meta_corr_s.append({
                            'Metrik': m,
                            'Ø Standard': np.mean(s_std_l),
                            'Ø Fair': np.mean(s_fair_l),
                            'Ø Sehr Fair': np.mean(s_sfair_l)
                        })
                        meta_corr_k.append({
                            'Metrik': m,
                            'Ø Standard': np.mean(k_std_l),
                            'Ø Fair': np.mean(k_fair_l),
                            'Ø Sehr Fair': np.mean(k_sfair_l)
                        })

                # B) WEIGHTED SCORES
                for (pid, playlist), group in analysis_df.groupby(['participant', 'playlist']):
                    raw_vals = group[m].values
                    subj_ranks = group['subj_rank'].values
                    vals, current_thresh = apply_tolerance_logic(raw_vals)
                    
                    # --- RANG BERECHNUNG ---
                    if "Standard" in global_ws_mode:
                        calc_ranks = calculate_physio_ranks(vals, 0.0) # Threshold 0
                        
                    elif "Sehr Fair" in global_ws_mode:
                        calc_ranks = np.zeros(len(vals))
                        sort_idx = np.argsort(vals)[::-1]
                        sorted_values = vals[sort_idx]
                        
                        current_rank_counter = 1
                        i_idx = 0
                        while i_idx < len(vals):
                            cluster_indices = [sort_idx[i_idx]]
                            j = i_idx + 1
                            while j < len(vals) and abs(sorted_values[i_idx] - sorted_values[j]) < current_thresh:
                                cluster_indices.append(sort_idx[j])
                                j += 1
                            
                            available_ranks = list(range(current_rank_counter, current_rank_counter + len(cluster_indices)))
                            cluster_subj_ranks = subj_ranks[cluster_indices]
                            sorted_rel_indices = np.argsort(cluster_subj_ranks)
                            
                            for k_idx, rel_idx in enumerate(sorted_rel_indices):
                                calc_ranks[cluster_indices[rel_idx]] = available_ranks[k_idx]
                                
                            current_rank_counter += len(cluster_indices)
                            i_idx = j
                    
                    else: # Fair
                        calc_ranks = calculate_physio_ranks(vals, current_thresh)

                    # --- SCORE BERECHNUNG ---
                    sorted_vals_desc = np.sort(vals)[::-1]
                    temp_grp = group.copy().reset_index(drop=True)
                    temp_grp['val_transformed'] = vals
                    temp_grp['calc_rank'] = calc_ranks # Hier nutzen wir den dynamischen Rang
                    
                    for idx, row in temp_grp.iterrows():
                        s_rank = int(row['subj_rank'])
                        c_rank = row['calc_rank']
                        val_act = row['val_transformed']
                        
                        ref_idx = min(s_rank - 1, len(sorted_vals_desc) - 1)
                        val_ref = sorted_vals_desc[ref_idx]
                        
                        rank_diff = abs(s_rank - c_rank)
                        val_diff = abs(val_act - val_ref)
                        
                        is_physio_error = val_diff >= current_thresh
                        
                        w_score = (rank_diff * val_diff) if is_physio_error else 0.0
                        
                        item_name = row['item']
                        specific_name = item_name
                        if item_name == 'Grundaktivität':
                            meta_weighted_scores.append({'Metrik': m, 'Song': 'Grundaktivität (Gesamt)', 'Weighted Score': w_score})
                            pl_raw = row['playlist'].lower()
                            if 'audio1' in pl_raw: suffix = 'Audio 1'
                            elif 'audio2' in pl_raw: suffix = 'Audio 2'
                            elif 'audio3' in pl_raw: suffix = 'Audio 3'
                            elif 'audio4' in pl_raw: suffix = 'Audio 4'
                            elif 'youtube' in pl_raw: suffix = 'Youtube'
                            else: suffix = pl_raw
                            specific_name = f"Grundaktivität ({suffix})"
                        
                        meta_weighted_scores.append({'Metrik': m, 'Song': specific_name, 'Weighted Score': w_score})

            st.subheader("Dominierender Entspannungstypus")
            
            tab_meta_s, tab_meta_k = st.tabs(["Spearman (Rang-Abstände)", "Kendall (Paar-Reihenfolge)"])
            
            with tab_meta_s:
                if meta_corr_s:
                    df_s = pd.DataFrame(meta_corr_s).set_index('Metrik').sort_values('Ø Sehr Fair', ascending=False)
                    st.table(df_s.style.background_gradient(cmap='RdYlGn', vmin=-1, vmax=1).format("{:.3f}"))
            
            with tab_meta_k:
                if meta_corr_k:
                    df_k = pd.DataFrame(meta_corr_k).set_index('Metrik').sort_values('Ø Sehr Fair', ascending=False)
                    st.table(df_k.style.background_gradient(cmap='RdYlGn', vmin=-1, vmax=1).format("{:.3f}"))

            st.subheader(f"Globale Einschätzungs-Schwierigkeit (Item-Analyse: {global_ws_mode})")
            if meta_weighted_scores:
                df_meta_ws = pd.DataFrame(meta_weighted_scores)
                item_difficulty = df_meta_ws.groupby('Song')['Weighted Score'].mean().reset_index()
                item_difficulty.columns = ['Song / Zustand', 'Ø Globaler Fehler-Score']
                item_difficulty = item_difficulty.sort_values(by='Ø Globaler Fehler-Score', ascending=False).reset_index(drop=True)
                
                c1, c2 = st.columns([2, 1])
                with c1: st.dataframe(item_difficulty.style.background_gradient(subset=['Ø Globaler Fehler-Score'], cmap='Reds').format({'Ø Globaler Fehler-Score': '{:.4f}'}))
                with c2:
                    if not item_difficulty.empty:
                        st.markdown(f"**Highlights:**\n🔴 **Schwierig:** *{item_difficulty.iloc[0]['Song / Zustand']}*\n🟢 **Klar:** *{item_difficulty.iloc[-1]['Song / Zustand']}*")

            # -----------------------------------
            # GENRE ANALYSE
            # -----------------------------------
            st.markdown("---")
            st.header("Genre-Analyse: Die „Entspannungs-Meisterschaft“")
            
            st.info(
                "Dieser Bereich vergleicht, welches Musik-Genre **subjektiv** (nach Meinung) vs. **objektiv** (nach EEG) "
                "am besten abgeschnitten hat. Durch ein **intra-subjektives Ranking** werden Unterschiede in der Signalstärke (z.B. dicker Schädel) herausgerechnet. **[Automatische Normalisierung durch Ränge]**"
            )

            st.subheader("Genre-Definition (Mapping)")
            
            raw_unique_songs = analysis_df[~analysis_df['item'].str.contains('Grundaktivität', case=False, na=False)]['item'].unique()
            unique_songs = sorted(raw_unique_songs, key=natural_sort_key)
            
            genre_mapping = {}
            
            with st.expander("🎵 Zuordnung Songs zu Genres", expanded=False):
                cols = st.columns(4)
                for i, song in enumerate(unique_songs):
                    with cols[i % 4]:
                        default_val = DEFAULT_GENRES.get(song, "")
                        val = st.text_input(f"{song}", value=default_val, placeholder="Genre...", key=f"genre_{i}")
                        genre_mapping[song] = val.strip().title() if val.strip() != "" else "Unbekannt"
            
            genre_df = analysis_df.copy()
            genre_df['genre'] = genre_df['item'].apply(lambda x: 'Ruhephase' if 'Grundaktivität' in x else genre_mapping.get(x, 'Unbekannt'))
            
            target_metric_genre = st.selectbox(
                "Wähle die physiologische Metrik für das EEG-Leaderboard:", 
                METRICS, index=1, 
                help="Welcher physiologische Zustand soll als 'Sieger' definiert werden?"
            )

            grp_genre = genre_df.groupby(['participant', 'genre']).agg({
                'subj_rank': 'mean',
                target_metric_genre: 'mean'
            }).reset_index()

            grp_genre['rank_norm_subj'] = grp_genre.groupby('participant')['subj_rank'].rank(method='average', ascending=True)
            grp_genre['rank_norm_eeg_std'] = grp_genre.groupby('participant')[target_metric_genre].rank(method='average', ascending=False)
            grp_genre['rank_norm_eeg_fair'] = np.nan 

            for pid in grp_genre['participant'].unique():
                idx = grp_genre['participant'] == pid
                subset = grp_genre[idx]
                raw_vals = subset[target_metric_genre].values
                vals_calc, current_thresh = apply_tolerance_logic(raw_vals)
                fair_ranks = calculate_physio_ranks(vals_calc, current_thresh)
                grp_genre.loc[idx, 'rank_norm_eeg_fair'] = fair_ranks

            leaderboard = grp_genre.groupby('genre').agg({
                'rank_norm_subj': 'mean',
                'rank_norm_eeg_std': 'mean',
                'rank_norm_eeg_fair': 'mean'
            }).reset_index()
            
            leaderboard.columns = ['Genre', 'Ø Rang (Gefühl)', 'Ø Rang (EEG Standard)', 'Ø Rang (EEG Fair)']

            def make_display_table(df, sort_col, val_col):
                d = df.sort_values(sort_col).reset_index(drop=True)
                d.index = d.index + 1
                d.index.name = 'Platz'
                return d[['Genre', val_col]]

            st.subheader("Analyse A: Die Erwartung (Subjektives Ranking)")
            
            with st.expander("📘 Erklärung: Das Glaubens-Ranking", expanded=False):
                st.markdown(
                    "**Fragestellung:** Welches Genre *glaubten* die Probanden, sei am entspannendsten?\n"
                    "**Metrik:** Durchschnittlicher Rangplatz (1 = Bester)."
                )
                
            c1, c2 = st.columns([1, 1])
            with c1:
                st.markdown("**Leaderboard (Gefühl):**")
                disp_subj = make_display_table(leaderboard, 'Ø Rang (Gefühl)', 'Ø Rang (Gefühl)')
                st.dataframe(
                    disp_subj.style
                    .background_gradient(cmap='Greens_r', subset=['Ø Rang (Gefühl)'])
                    .format("{:.2f}", subset=['Ø Rang (Gefühl)']) 
                )
            with c2:
                df_subj_sorted = leaderboard.sort_values('Ø Rang (Gefühl)')
                chart_s = alt.Chart(df_subj_sorted).mark_bar().encode(
                    x=alt.X('Ø Rang (Gefühl)', title='Ø Platzierung (1=Top)'),
                    y=alt.Y('Genre', sort=alt.EncodingSortField(field="Ø Rang (Gefühl)", order="ascending")),
                    color=alt.Color('Ø Rang (Gefühl)', scale=alt.Scale(scheme='greens', reverse=True), legend=None),
                    tooltip=['Genre', alt.Tooltip('Ø Rang (Gefühl)', format='.2f')]
                )
                st.altair_chart(chart_s, width='stretch')

            st.subheader("Analyse B: Die physiologische Realität (EEG Ranking)")

            st.info(
                    "ℹ️ **Methodik-Hinweis: Warum 'Fair' und nicht 'Sehr Fair'?**\n\n"
                    "In dieser Analyse wollen wir den **objektiven Sieger** ermitteln. \n"
                    "* **Standard:** Ignoriert Messrauschen. Zufällige Schwankungen entscheiden über Platz 1 oder 2.\n"
                    "* **Fair:** Betrachtet physiologisch gleiche Werte als **Unentschieden** (z.B. beide Platz 1.5). Das glättet das Leaderboard.\n"
                    "* **Warum nicht 'Sehr Fair'?** Die 'Sehr Fair'-Logik würde das Ranking manipulieren, um es dem *User-Geschmack* anzupassen. "
                    "Das wäre hier wissenschaftlich unsauber, da wir ja messen wollen, was der Körper *unabhängig* von der Meinung sagt."
                )
            
            col_std, col_fair = st.columns(2)
            
            with col_std:
                st.markdown("### Standard")
                st.caption("Mathematisch striktes Ranking")
                disp_std = make_display_table(leaderboard, 'Ø Rang (EEG Standard)', 'Ø Rang (EEG Standard)')
                st.dataframe(
                    disp_std.style
                    .background_gradient(cmap='Blues_r', subset=['Ø Rang (EEG Standard)'])
                    .format("{:.2f}", subset=['Ø Rang (EEG Standard)'])
                )
                df_std_sorted = leaderboard.sort_values('Ø Rang (EEG Standard)')
                chart_std = alt.Chart(df_std_sorted).mark_bar().encode(
                    x=alt.X('Ø Rang (EEG Standard)', title='Ø Platz (Standard)'),
                    y=alt.Y('Genre', sort=alt.EncodingSortField(field="Ø Rang (EEG Standard)", order="ascending")),
                    color=alt.Color('Ø Rang (EEG Standard)', scale=alt.Scale(scheme='blues', reverse=True), legend=None),
                    tooltip=['Genre', alt.Tooltip('Ø Rang (EEG Standard)', format='.2f')]
                ).properties(height=300)
                st.altair_chart(chart_std, width='stretch')

            with col_fair:
                st.markdown("### Fair ")
                st.caption("Physiologisch bereinigtes Ranking")
                disp_fair = make_display_table(leaderboard, 'Ø Rang (EEG Fair)', 'Ø Rang (EEG Fair)')
                st.dataframe(
                    disp_fair.style
                    .background_gradient(cmap='Blues_r', subset=['Ø Rang (EEG Fair)'])
                    .format("{:.2f}", subset=['Ø Rang (EEG Fair)'])
                )
                df_fair_sorted = leaderboard.sort_values('Ø Rang (EEG Fair)')
                chart_fair = alt.Chart(df_fair_sorted).mark_bar().encode(
                    x=alt.X('Ø Rang (EEG Fair)', title='Ø Platz (Fair)'),
                    y=alt.Y('Genre', sort=alt.EncodingSortField(field="Ø Rang (EEG Fair)", order="ascending")),
                    color=alt.Color('Ø Rang (EEG Fair)', scale=alt.Scale(scheme='blues', reverse=True), legend=None),
                    tooltip=['Genre', alt.Tooltip('Ø Rang (EEG Fair)', format='.2f')]
                ).properties(height=300)
                st.altair_chart(chart_fair, width='stretch')

            st.markdown("### 🔍 Vergleich: Wahrnehmung vs. Realität")
            
            st.info("Hier sehen wir, ob die Teilnehmer ein Genre besser oder schlechter einschätzen, als es ihr Körper tatsächlich empfindet. Positive Werte bedeuten: Der Körper war entspannter, als der Kopf dachte.")

            leaderboard['Delta (Std)'] = leaderboard['Ø Rang (Gefühl)'] - leaderboard['Ø Rang (EEG Standard)']
            leaderboard['Fazit (Std)'] = leaderboard['Delta (Std)'].apply(
                lambda x: "Besser als gedacht" if x > 0.5 else ("Schlechter als gedacht" if x < -0.5 else "Treffer ✅")
            )

            leaderboard['Delta (Fair)'] = leaderboard['Ø Rang (Gefühl)'] - leaderboard['Ø Rang (EEG Fair)']
            leaderboard['Fazit (Fair)'] = leaderboard['Delta (Fair)'].apply(
                lambda x: "Besser als gedacht" if x > 0.5 else ("Schlechter als gedacht" if x < -0.5 else "Treffer ✅")
            )
            
            f_col1, f_col2 = st.columns(2)

            with f_col1:
                st.markdown("**A) Standard**")
                disp_std = leaderboard.sort_values('Delta (Std)', ascending=False).reset_index(drop=True)
                disp_std.index = disp_std.index + 1
                disp_std.index.name = 'Platz'
                st.dataframe(
                    disp_std[['Genre', 'Ø Rang (Gefühl)', 'Ø Rang (EEG Standard)', 'Delta (Std)', 'Fazit (Std)']]
                    .style
                    .background_gradient(subset=['Delta (Std)'], cmap='RdYlGn', vmin=-2, vmax=2)
                    .format("{:.2f}", subset=['Ø Rang (Gefühl)', 'Ø Rang (EEG Standard)', 'Delta (Std)'])
                )

            with f_col2:
                st.markdown("**B) Fair**")
                disp_fair = leaderboard.sort_values('Delta (Fair)', ascending=False).reset_index(drop=True)
                disp_fair.index = disp_fair.index + 1
                disp_fair.index.name = 'Platz'
                st.dataframe(
                    disp_fair[['Genre', 'Ø Rang (Gefühl)', 'Ø Rang (EEG Fair)', 'Delta (Fair)', 'Fazit (Fair)']]
                    .style
                    .background_gradient(subset=['Delta (Fair)'], cmap='RdYlGn', vmin=-2, vmax=2)
                    .format("{:.2f}", subset=['Ø Rang (Gefühl)', 'Ø Rang (EEG Fair)', 'Delta (Fair)'])
                )

            # -----------------------------------
            # POLARISIERUNGS-MATRIX
            # -----------------------------------
            st.subheader("Universalität vs. Individualität (Polarisierungs-Matrix)")

            with st.expander("📘 Warum hier immer Z-Score", expanded=False):
                st.markdown(
                    """
                    In den vorherigen Rankings (6.2/6.3) ging es um **Ränge**. Hier schauen wir auf die **Signalstärke**.
                    
                    **Das Problem:** Proband A hat EEG-Werte um 5 µV, Proband B um 50 µV. Würden wir das einfach mischen, 
                    würde Proband B die Statistik dominieren.
                    
                    **Die Lösung (Z-Score):** Wir normalisieren die Daten. 
                    * **0.00** = Der Durchschnittszustand des Probanden.
                    * **+1.00** = Deutlich entspannter als sonst (1 Standardabweichung über Schnitt).
                    * **-1.00** = Deutlich weniger entspannt als sonst.
                    
                    So können wir die **Wirkung der Musik** vergleichen, unabhängig von dem Session Kontext.
                    """
                )

            z_data_list = []
            
            for pid, group in genre_df.groupby('participant'):
                raw_vals = group[target_metric_genre].values
                std_dev = np.std(raw_vals)
                if std_dev == 0: 
                    z_scores = np.zeros(len(raw_vals))
                else:
                    z_scores = (raw_vals - np.mean(raw_vals)) / std_dev
                temp_df = group.copy()
                temp_df['z_score'] = z_scores
                z_data_list.append(temp_df)
                
            df_z_matrix = pd.concat(z_data_list)
            
            matrix_stats = df_z_matrix.groupby('genre')['z_score'].agg(['mean', 'std', 'count']).reset_index()
            matrix_stats.columns = ['Genre', 'Effektivität', 'Individualität', 'N']
            
            matrix_stats['Effektivität'] = matrix_stats['Effektivität'].round(3)
            matrix_stats['Individualität'] = matrix_stats['Individualität'].round(3)

            with st.expander("📊 Graph", expanded=False):
                base = alt.Chart(matrix_stats).encode(
                    x=alt.X('Effektivität:Q', title='← Weniger Wirkung ... Globale Wirkung (Ø Z-Score) ... Mehr Wirkung →'),
                    y=alt.Y('Individualität:Q', title='Konsens ... Individuelle Varianz (Std) ... Polarisierung →'),
                    tooltip=['Genre', 'Effektivität', 'Individualität', 'N']
                )
                
                vline = alt.Chart(pd.DataFrame({'x': [0]})).mark_rule(strokeDash=[5,5], color='gray').encode(x='x')
                
                median_std = matrix_stats['Individualität'].median()
                hline = alt.Chart(pd.DataFrame({'y': [median_std]})).mark_rule(strokeDash=[5,5], color='gray').encode(y='y')
                
                points = base.mark_circle(size=300, opacity=0.8).encode(
                    color=alt.Color('Effektivität:Q', scale=alt.Scale(scheme='redyellowgreen', domain=[-0.5, 0.5]), legend=None)
                )
                
                text = base.mark_text(align='left', baseline='middle', dx=15, fontWeight='bold').encode(
                    text='Genre'
                )
                
                chart_matrix = (points + text + vline + hline).properties(height=500).interactive()
                st.altair_chart(chart_matrix, width='stretch')
            
            c1, c2 = st.columns([1, 1])
            with c1:
                st.markdown("##### Datentabelle")
                st.dataframe(
                    matrix_stats.sort_values('Effektivität', ascending=False)
                    .style
                    .background_gradient(subset=['Effektivität'], cmap='RdYlGn', vmin=-0.5, vmax=0.5)
                    .background_gradient(subset=['Individualität'], cmap='Reds')
                    .format("{:.2f}", subset=['Effektivität', 'Individualität'])
                )
            with c2:
                st.markdown("##### 💡 Lesehilfe für die Zahlen")
                st.markdown("**1. Spalte: Effektivität (Ø Z-Score)**")
                st.caption("Gibt an, wie stark das Genre im Vergleich zum persönlichen Durchschnitt wirkt.")
                st.markdown(
                    """
                    * `> 0.00`: Das Genre wirkt **überdurchschnittlich** gut.
                    * `0.00`: Durchschnittliche Wirkung.
                    * `< 0.00`: Das Genre wirkt weniger gut als der Rest.
                    """
                )
                st.markdown("---")
                st.markdown("**2. Spalte: Individualität (Std-Abweichung)**")
                st.caption("Gibt an, wie sehr sich die Probanden bei diesem Genre 'uneinig' sind.")
                st.markdown(
                    """
                    * **Niedriger Wert (0.0 - 0.5):** Hoher Konsens. Das Genre wirkt auf fast alle gleich. (Sicherer Tipp).
                    * **Hoher Wert (> 0.8):** Starke Polarisierung. Manche lieben es physiologisch, andere gar nicht. (Riskanter Tipp).
                    """
                )

        # ==========================================================================================
        # TAB 3: META-DATEN ANALYSE
        # ==========================================================================================
        with tab_meta:
            st.header("Meta Daten Analyse")
            
            df_7 = df_full[~df_full['playlist'].str.contains('youtube', case=False)].copy()
            
            if 'item_entspannt' not in df_7.columns:
                st.error("Die Spalte 'item_entspannt' wurde in den Daten nicht gefunden.")
            else:
                st.subheader("Identifikation von Widersprüchen (Physiologisch validiert)")
                
                with st.expander("📘 Methodik: Wann gilt ein Widerspruch als 'echt'?", expanded=False):
                    st.markdown(
                        """
                        Hier prüfen wir, ob die Messdaten(EEG-Ranking) der Aussage des Probanden ("Entspannt mich: JA/NEIN") widersprechen.
                        
                        **Standard:**
                        * Sobald der Messwert bei "Nein" höher ist als bei "Ja" ($Val_{Nein} > Val_{Ja}$), gilt dies als Widerspruch.
                        
                        **Fair:**
                        * Ein Widerspruch gilt erst dann als echt, wenn die Differenz den **Toleranz-Schwellenwert** überschreitet ($Val_{Nein} - Val_{Ja} > Threshold$).
                        """
                    )
                
                c_meta1, c_meta2 = st.columns(2)
                with c_meta1:
                    target_metric_7 = st.selectbox(
                        "Vergleichs-Metrik wählen:", 
                        METRICS, 
                        index=1,
                        key="metric_7"
                    )
                with c_meta2:
                    meta_mode = st.selectbox(
                        "Analyse-Modus wählen:",
                        ["Standard", "Fair"],
                        index=1, # Default auf Fair
                        key="meta_mode_selector"
                    )
                
                df_7_relevant = df_7[df_7['item_entspannt'].isin(['Ja', 'Nein'])].copy()
                
                paradox_cases = []
                analyzed_playlists_count = 0
                
                for (pid, playlist), group in df_7_relevant.groupby(['participant', 'playlist']):
                    items_yes = group[group['item_entspannt'] == 'Ja']
                    items_no = group[group['item_entspannt'] == 'Nein']
                    
                    if items_yes.empty or items_no.empty:
                        continue
                    
                    analyzed_playlists_count += 1
                    
                    raw_vals_playlist = group[target_metric_7].values
                    _, current_thresh = apply_tolerance_logic(raw_vals_playlist)
                    
                    comparison_threshold = current_thresh if "Fair" in meta_mode else 0.0
                    
                    for idx_y, row_yes in items_yes.iterrows():
                        for idx_n, row_no in items_no.iterrows():
                            val_yes = row_yes[target_metric_7]
                            val_no = row_no[target_metric_7]
                            
                            diff = val_no - val_yes
                            
                            if diff > comparison_threshold:
                                paradox_cases.append({
                                    'Proband': pid,
                                    'Playlist': playlist,
                                    'Item (User: JA)': row_yes['item'],
                                    'Item (User: NEIN)': row_no['item'],
                                    f'Wert (JA)': val_yes,
                                    f'Wert (NEIN)': val_no,
                                    'Differenz': diff,
                                    'Threshold (Ref)': current_thresh
                                })
                
                if paradox_cases:
                    df_paradox = pd.DataFrame(paradox_cases)
                    df_paradox = df_paradox.sort_values(by='Differenz', ascending=False).reset_index(drop=True)
                    
                    st.markdown("#### A) Echte physiologische Widersprüche (Inkl. Grundaktivität)")
                    
                    if analyzed_playlists_count > 0:
                        total_contradictions = len(df_paradox)
                        avg_contradictions = total_contradictions / analyzed_playlists_count
                        n_users_paradox = df_7_relevant['participant'].nunique()
                        
                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric("Analysierte Proband", n_users_paradox)
                        c2.metric("Analysierte Playlists", analyzed_playlists_count)
                        c3.metric("Gefundene Widersprüche", total_contradictions)
                        c4.metric("Ø pro Playlist", f"{avg_contradictions:.1f}")
                        
                    with st.expander("Details", expanded=False):
                        st.dataframe(
                            df_paradox.style
                            .background_gradient(subset=['Differenz'], cmap='Reds')
                            .format({'Wert (JA)': '{:.4f}', 'Wert (NEIN)': '{:.4f}', 'Differenz': '{:.4f}', 'Threshold (Ref)': '{:.4f}'})
                        )
                    
                    st.markdown("#### B) Nur Song-Widersprüche (Ohne Grundaktivität)")
                    
                    mask_songs_only = (
                        ~df_paradox['Item (User: JA)'].str.contains('Grundaktivität', case=False) & 
                        ~df_paradox['Item (User: NEIN)'].str.contains('Grundaktivität', case=False)
                    )
                    df_paradox_songs = df_paradox[mask_songs_only].reset_index(drop=True)
                    
                    if analyzed_playlists_count > 0:
                        total_songs_contradictions = len(df_paradox_songs)
                        avg_songs_contradictions = total_songs_contradictions / analyzed_playlists_count
                        
                        c1, c2, c3, c4 = st.columns(4)
                        c3.metric("Gefundene Widersprüche", total_songs_contradictions)
                        c4.metric("Ø pro Playlist", f"{avg_songs_contradictions:.1f}")

                    if not df_paradox_songs.empty:
                        with st.expander("Details", expanded=False):
                            st.dataframe(
                                df_paradox_songs.style
                                .background_gradient(subset=['Differenz'], cmap='Reds')
                                .format({'Wert (JA)': '{:.4f}', 'Wert (NEIN)': '{:.4f}', 'Differenz': '{:.4f}', 'Threshold (Ref)': '{:.4f}'})
                            )
                    else:
                        st.success(f"Im Modus '{meta_mode}' gibt es keine Widersprüche zwischen Songs.")

                else:
                    if analyzed_playlists_count > 0:
                        st.success(f"Keine Widersprüche im Modus '{meta_mode}' gefunden!")
                    else:
                        st.warning("Keine analysierbaren Daten (zu wenig Ja/Nein Angaben).")

                # -----------------------------------
                # SUBJEKTIVE TENDENZEN
                # -----------------------------------
                st.markdown("---")
                st.subheader("Subjektive Tendenzen: Geschmack vs. Entspannung")
                
                st.info(
                    "Hier prüfen wir das Antwortverhalten rein auf subjektiver Ebene (ohne EEG).\n\n"
                )
                
                df_72 = df_7[~df_7['item'].str.contains('Grundaktivität', case=False, na=False)].copy()
                df_72 = df_72.dropna(subset=['item_entspannt', 'item_gefallen', 'subj_rank'])
                
                if df_72.empty:
                    st.warning("Keine ausreichenden Daten für diese Analyse.")
                else:
                    st.markdown("#### A) Zusammenhang: Gefallen ↔ Entspannen")
                    c1, c2 = st.columns(2)
                    
                    with c1:
                        relaxed_songs = df_72[df_72['item_entspannt'] == 'Ja']
                        count_relaxed = len(relaxed_songs)
                        if count_relaxed > 0:
                            counts_g = relaxed_songs['item_gefallen'].value_counts()
                            perc_ja = (counts_g.get('Ja', 0) / count_relaxed) * 100
                            st.metric("Ist Entspannt → Gefällt mir", f"{perc_ja:.1f}%")
                        else:
                            st.metric("Ist Entspannt → Gefällt mir", "N/A")

                    with c2:
                        liked_songs = df_72[df_72['item_gefallen'] == 'Ja']
                        count_liked = len(liked_songs)
                        if count_liked > 0:
                            counts_e = liked_songs['item_entspannt'].value_counts()
                            perc_relaxed = (counts_e.get('Ja', 0) / count_liked) * 100
                            st.metric("Gefällt mir → Ist Entspannt", f"{perc_relaxed:.1f}%")
                        else:
                             st.metric("Gefällt mir → Ist Entspannt", "N/A")

                    st.markdown("---")
                    st.markdown("#### B) Die 'Ranking-Anomalie' (Unbeliebt schlägt Beliebt)")
                    st.caption("Wie oft gewinnt ein Song aus Kategorie X gegen einen 'Gefällt mir'-Song (Ja)?")
                    
                    stats_b = {
                        'all': {'comps': 0, 'anoms': 0},      
                        'neutral': {'comps': 0, 'anoms': 0}, 
                        'nein': {'comps': 0, 'anoms': 0}     
                    }
                    
                    for (pid, pl), group in df_72.groupby(['participant', 'playlist']):
                        group_liked = group[group['item_gefallen'] == 'Ja']
                        if group_liked.empty: continue

                        for _, row_liked in group_liked.iterrows():
                            for _, row_other in group.iterrows():
                                cat_other = row_other['item_gefallen']
                                if cat_other == 'Ja': continue
                                
                                is_anomaly = row_other['subj_rank'] < row_liked['subj_rank']
                                
                                stats_b['all']['comps'] += 1
                                if is_anomaly: stats_b['all']['anoms'] += 1
                                
                                if cat_other == 'Neutral':
                                    stats_b['neutral']['comps'] += 1
                                    if is_anomaly: stats_b['neutral']['anoms'] += 1
                                    
                                if cat_other == 'Nein':
                                    stats_b['nein']['comps'] += 1
                                    if is_anomaly: stats_b['nein']['anoms'] += 1

                    c1, c2, c3 = st.columns(3)
                    
                    def show_anomaly_metric(col, label, data):
                        if data['comps'] > 0:
                            perc = (data['anoms'] / data['comps']) * 100
                            col.metric(
                                label=label,
                                value=f"{perc:.1f}%",
                                delta=f"{data['anoms']} / {data['comps']} Fälle",
                                delta_color="inverse"
                            )
                        else:
                            col.metric(label, "Keine Daten")

                    show_anomaly_metric(c1, "Gegen 'Neutral' + 'Nein'", stats_b['all'])
                    show_anomaly_metric(c2, "Nur gegen 'Neutral'", stats_b['neutral'])
                    show_anomaly_metric(c3, "Nur gegen 'Nein'", stats_b['nein'])

                    st.markdown("---")
                    st.markdown("#### C) 'Gefällt mir'-Quote der Top-Plätze")
                    st.caption("Wie oft entsprach der Song auf Platz 1, 2 oder 3 dem Geschmack?")

                    top_stats = []
                    for (pid, pl), group in df_72.groupby(['participant', 'playlist']):
                        for rank in [1, 2, 3]:
                            row = group[group['subj_rank'] == rank]
                            if not row.empty:
                                val = row.iloc[0]['item_gefallen']
                                is_ja = (val == 'Ja')
                                is_ja_neutral = (val in ['Ja', 'Neutral'])
                                top_stats.append({
                                    'Rang': rank,
                                    'Gefallen (Ja)': is_ja,
                                    'Gefallen (Ja + Neutral)': is_ja_neutral
                                })
                    
                    if top_stats:
                        df_top = pd.DataFrame(top_stats)
                        summary = df_top.groupby('Rang').agg({
                            'Gefallen (Ja)': 'mean',
                            'Gefallen (Ja + Neutral)': 'mean'
                        }).reset_index()
                        
                        summary['Gefallen (Ja)'] = (summary['Gefallen (Ja)'] * 100).map('{:.1f}%'.format)
                        summary['Gefallen (Ja + Neutral)'] = (summary['Gefallen (Ja + Neutral)'] * 100).map('{:.1f}%'.format)
                        
                        st.dataframe(
                            summary.set_index('Rang').style.background_gradient(cmap='Greens')
                        )
                    else:
                        st.info("Keine Ranking-Daten verfügbar.")

                    st.markdown("---")
                    st.caption("Detail-Ansicht: Durchschnittlicher Rangplatz pro Kategorie")
                    
                    rank_stats = df_72.groupby('item_gefallen')['subj_rank'].mean().reset_index()
                    sorter = {'Ja': 0, 'Neutral': 1, 'Nein': 2}
                    rank_stats['sort_key'] = rank_stats['item_gefallen'].map(sorter)
                    rank_stats = rank_stats.sort_values('sort_key')
                    
                    chart_ranks = alt.Chart(rank_stats).mark_bar().encode(
                        x=alt.X('item_gefallen', title='Gefallen?', sort=['Ja', 'Neutral', 'Nein']),
                        y=alt.Y('subj_rank', title='Ø Platzierung (1 = Bester)'),
                        color=alt.Color('item_gefallen', legend=None, scale=alt.Scale(domain=['Ja', 'Neutral', 'Nein'], range=['#2ECC71', '#F1C40F', '#E74C3C'])),
                        tooltip=['item_gefallen', alt.Tooltip('subj_rank', format='.2f')]
                    ).properties(height=250)
                    st.altair_chart(chart_ranks, width='stretch')

        # ==========================================================================================
        # TAB 4: VERLAUFSANALYSE
        # ==========================================================================================
        with tab_verlauf:
            # -----------------------------------
            # LERN-EFFEKT ANALYSE
            # -----------------------------------
            st.header("Lern-Effekt: Verbessert sich die Einschätzung? (Differenz der Korrelationskoeffizienten)")
            
            c9_1, c9_2 = st.columns(2)
            
            with c9_1:
                target_metric_9 = st.selectbox(
                    "Basis-Metrik:", 
                    METRICS, 
                    index=1,
                    key="metric_9"
                )
                
            with c9_2:
                calc_mode_9 = st.selectbox(
                    "Berechnungs-Modus:", 
                    ["Standard", "Fair", "Sehr Fair"], 
                    index=1, # Default: Fair
                    key="mode_9"
                )

            df_9 = df_full[~df_full['playlist'].str.contains('youtube', case=False)].copy()
            
            def get_learning_session_id(pl_name):
                pl = pl_name.lower()
                if any(x in pl for x in ['audio1', 'audio2']): return 'S1'
                if any(x in pl for x in ['audio3', 'audio4']): return 'S2'
                return None

            df_9['session_type'] = df_9['playlist'].apply(get_learning_session_id)
            df_9 = df_9.dropna(subset=['session_type'])

            learning_stats_s = [] 
            learning_stats_k = [] 
            
            for pid in df_9['participant'].unique():
                p_data = df_9[df_9['participant'] == pid]
                
                s1_data = p_data[p_data['session_type'] == 'S1'].copy()
                s2_data = p_data[p_data['session_type'] == 'S2'].copy()
                
                if s1_data.empty or s2_data.empty: continue
                
                # --- SESSION 1 ---
                vals_s1, thresh_s1 = apply_tolerance_logic(s1_data[target_metric_9].values)
                s1_data[target_metric_9] = vals_s1
                (s_std1, s_fair1, s_sfair1), (k_std1, k_fair1, k_sfair1) = calculate_correlation_variants(s1_data, target_metric_9, 'subj_rank', thresh_s1)
                
                # --- SESSION 2 ---
                vals_s2, thresh_s2 = apply_tolerance_logic(s2_data[target_metric_9].values)
                s2_data[target_metric_9] = vals_s2 
                (s_std2, s_fair2, s_sfair2), (k_std2, k_fair2, k_sfair2) = calculate_correlation_variants(s2_data, target_metric_9, 'subj_rank', thresh_s2)
                
                learning_stats_s.append({
                    'Proband': pid,
                    'S1_Standard': s_std1,   'S2_Standard': s_std2,   'Δ Standard': s_std2 - s_std1,
                    'S1_Fair': s_fair1,      'S2_Fair': s_fair2,      'Δ Fair': s_fair2 - s_fair1,
                    'S1_SehrFair': s_sfair1, 'S2_SehrFair': s_sfair2, 'Δ Sehr Fair': s_sfair2 - s_sfair1
                })
                learning_stats_k.append({
                    'Proband': pid,
                    'S1_Standard': k_std1,   'S2_Standard': k_std2,   'Δ Standard': k_std2 - k_std1,
                    'S1_Fair': k_fair1,      'S2_Fair': k_fair2,      'Δ Fair': k_fair2 - k_fair1,
                    'S1_SehrFair': k_sfair1, 'S2_SehrFair': k_sfair2, 'Δ Sehr Fair': k_sfair2 - k_sfair1
                })

            if not learning_stats_s:
                st.warning("Keine ausreichenden Daten (Stelle sicher, dass Audio 1/2 und Audio 3/4 vorhanden sind).")
            else:
                col_map = {
                    "Standard":  ('S1_Standard', 'S2_Standard', 'Δ Standard'),
                    "Fair":      ('S1_Fair', 'S2_Fair', 'Δ Fair'),
                    "Sehr Fair": ('S1_SehrFair', 'S2_SehrFair', 'Δ Sehr Fair')
                }
                
                col_s1, col_s2, col_delta = col_map[calc_mode_9]
                numeric_cols = ['Δ Standard', 'Δ Fair', 'Δ Sehr Fair']

                tab_s, tab_k = st.tabs(["Spearman", "Kendall"])
                
                with tab_s:
                    df_perf_s = pd.DataFrame(learning_stats_s)
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        avg_s1 = df_perf_s[col_s1].mean()
                        st.metric(f"Ø Genauigkeit S1 ({calc_mode_9})", f"{avg_s1:.3f}")
                    with c2:
                        avg_delta = df_perf_s[col_delta].mean()
                        st.metric(f"Ø Lerneffekt ({calc_mode_9})", f"{avg_delta:+.3f}", delta_color="normal" if avg_delta >= 0 else "inverse")

                    st.markdown(f"#### Individueller Trend ({calc_mode_9})")
                    
                    df_slope_s = df_perf_s.melt(
                        id_vars=['Proband'], 
                        value_vars=[col_s1, col_s2], 
                        var_name='Session_Raw', 
                        value_name='Korrelation'
                    )
                    df_slope_s['Session'] = df_slope_s['Session_Raw'].apply(lambda x: 'Session 1' if 'S1' in x else 'Session 2')

                    chart_s = alt.Chart(df_slope_s).mark_line(point=True).encode(
                        x=alt.X('Session', title=None),
                        y=alt.Y('Korrelation', scale=alt.Scale(domain=[-1, 1])),
                        color='Proband',
                        tooltip=['Proband', 'Session', alt.Tooltip('Korrelation', format='.3f')]
                    ).properties(height=300)
                    st.altair_chart(chart_s, width='stretch')

                    st.markdown("#### Detail-Tabelle (Deltas)")
                    st.dataframe(
                        df_perf_s[['Proband', 'Δ Standard', 'Δ Fair', 'Δ Sehr Fair']]
                        .sort_values(col_delta, ascending=False)
                        .style
                        .background_gradient(subset=numeric_cols, cmap='RdYlGn', vmin=-0.3, vmax=0.3)
                        .format("{:+.3f}", subset=numeric_cols) 
                    )

                with tab_k:
                    df_perf_k = pd.DataFrame(learning_stats_k)
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        avg_k1 = df_perf_k[col_s1].mean()
                        st.metric(f"Ø Genauigkeit S1 ({calc_mode_9})", f"{avg_k1:.3f}")
                    with c2:
                        avg_k_delta = df_perf_k[col_delta].mean()
                        st.metric(f"Ø Lerneffekt ({calc_mode_9})", f"{avg_k_delta:+.3f}", delta_color="normal" if avg_k_delta >= 0 else "inverse")

                    st.markdown(f"#### Individueller Trend ({calc_mode_9})")
                    
                    df_slope_k = df_perf_k.melt(
                        id_vars=['Proband'], 
                        value_vars=[col_s1, col_s2], 
                        var_name='Session_Raw', 
                        value_name='Korrelation'
                    )
                    df_slope_k['Session'] = df_slope_k['Session_Raw'].apply(lambda x: 'Session 1' if 'S1' in x else 'Session 2')

                    chart_k = alt.Chart(df_slope_k).mark_line(point=True).encode(
                        x=alt.X('Session', title=None),
                        y=alt.Y('Korrelation', scale=alt.Scale(domain=[-1, 1])),
                        color='Proband',
                        tooltip=['Proband', 'Session', alt.Tooltip('Korrelation', format='.3f')]
                    ).properties(height=300)
                    st.altair_chart(chart_k, width='stretch')

                    st.markdown("#### Detail-Tabelle (Deltas)")
                    st.dataframe(
                        df_perf_k[['Proband', 'Δ Standard', 'Δ Fair', 'Δ Sehr Fair']]
                        .sort_values(col_delta, ascending=False)
                        .style
                        .background_gradient(subset=numeric_cols, cmap='RdYlGn', vmin=-0.3, vmax=0.3)
                        .format("{:+.3f}", subset=numeric_cols)
                    )
            st.markdown("---")
            # -----------------------------------
            # SESSION-ANALYSE (RELIABILITÄT DER MUSIK EFEKTE)
            # -----------------------------------
            st.header("Session-Konsistenz: Effekt der Songs über Kontexte")
            
            st.info(
                    """
                    Wie stark hat der Kontextwechsel den Effekt der Songs beeinflusst?\n
                    * Anderer Vorgänger (Song).
                    * Anderer Tag (mind. 1 Woche nach 1ter Aufnahme).
                    * Andere Zusammenstellung der Lieder in einer Playlist.
                    """
            )

            with st.expander("📘 Methodik: Session-Pooling & Kontext-Analyse", expanded=False):
                st.markdown("### Das Setup")
                st.markdown(
                    """
                    Wir vergleichen hier zwei Aufnahmen, die **mindestens eine Woche auseinander** lagen.
                    Dabei wurden die gleichen Songs gehört, jedoch in **unterschiedlichen Playlists** (Zusammenstellungen) und **Reihenfolgen** (Shuffle).
                    * **Ziel:** Wir wollen herausfinden, wie stark der **Kontext** (Vorgänger-Song, Tagesform) die Wirkung eines Liedes beeinflusst.
                    * **Datenbasis:** Nur *Audio 1/2* (Session 1) und *Audio 3/4* (Session 2).
                    """
                )

                st.divider()

                st.markdown("### Das Problem: Der 'Shuffle-Effekt'")
                c_old, c_new = st.columns(2)

                with c_old:
                    st.error("🔴 **Warum Playlist-Normalisierung scheitert**")
                    st.markdown(
                        """
                        Das Hauptproblem ist die **Neu-Verteilung (Distribution)** der Songs auf die Playlists.
                        
                        * **Das Szenario:** Ein Song aus *Audio 1* (Session 1) landet in *Audio 4* (Session 2). Er befindet sich in einer völlig neuen Gruppe von Songs.
                        * **Die Konsequenz:** Da sich die **Zusammensetzung** der Playlist geändert hat, verschiebt sich deren Mittelwert. Ein lokaler Z-Score wäre nicht mehr vergleichbar.
                        * **Vergleich:** Hätten wir nur die *Reihenfolge* innerhalb der gleichen Playlist geändert, wäre der lokale Z-Score noch valide gewesen. Da wir aber die "Töpfe" neu gemischt haben, brauchen wir den globalen Tages-Durchschnitt als stabilen Anker.
                        """
                    )

                with c_new:
                    st.success("🟢 **Lösung: Globales Session-Pooling**")
                    st.markdown(
                        """
                        Wir werfen alle Songs (+bl) eines Tages in einen Topf, um den Maßstab zu bilden.
                        
                        * **Die Rechnung:** Wir berechnen den Z-Score.
                        * **Der Vorteil:** Wir korrigieren die Tagesform (z.B. Hautwiderstand nach 1 Woche), behalten aber die relative Stärke der Songs bei, egal in welcher Playlist sie landeten.
                        """
                    )
            # ---------------------------------------------------------
            # EXPANDER: BERECHNUNG & INTERPRETATION 
            # ---------------------------------------------------------
            with st.expander("📘 Methodik: Berechnung, Schwellenwerte & Interpretation", expanded=False):
                st.markdown(r"### 1. Das Delta ($\Delta$)")
                st.latex(r"\Delta = |Z_{Session1} - Z_{Session2}|")
                
                st.markdown("### 2. Woher kommen die Schwellenwerte?")
                st.markdown(
                    r"""
                    Wir arbeiten hier mit **Z-Scores** (Standardabweichungen). Die Schwellenwerte leiten sich aus der Statistik der Normalverteilung ab:
                    
                    * **$\Delta < 0.5$ (Rauschen):** Eine Verschiebung um eine halbe Standardabweichung ist physiologisch oft vernachlässigbar und liegt im Bereich der natürlichen Varianz.
                    * **$\Delta > 1.0$ (Signifikant):** Eine Verschiebung um **eine ganze Standardabweichung** bedeutet, dass der Song seine relative Position im Datensatz massiv verändert hat (z.B. vom Durchschnittswert in die Top 15% der entspanntesten Werte).
                    """
                )

                st.markdown("### 3. Was sagt das über den Kontext aus?")
                st.info(
                    """
                    **Die Logik:** Da der Song selbst (die Audio-Datei) identisch blieb, aber der **Kontext** (Vorgänger-Song, Playlist-Mix) verändert wurde, interpretieren wir das Delta als Maß für die **Kontext-Sensitivität**.
                    
                    * **Hohes Delta:** Die Wirkung des Songs ist nicht inhärent fest, sondern wird stark durch das "Priming" des vorherigen Songs oder Aufnahme Kontext bestimmt (z.B. wirkt er nur ruhig, wenn vorher Metal lief, oder an einem weniger stressigen Tag).
                    * **Niedriges Delta:** Der Song setzt sich physiologisch gegen den Kontext durch.
                    """
                )


            # ---------------------------------------------------------
            # BERECHNUNG: SESSION POOLING
            # ---------------------------------------------------------
            
            # 1. Daten filtern: Youtube raus, nur Audio Playlists rein
            df_cons = df_full[
                (~df_full['playlist'].str.contains('youtube', case=False)) 
            ].copy()

            # 2. Session ID zuweisen
            def assign_session(pl_name):
                pl = pl_name.lower()
                if 'audio1' in pl or 'audio2' in pl: return 'Session 1'
                if 'audio3' in pl or 'audio4' in pl: return 'Session 2'
                return None

            df_cons['Session_ID'] = df_cons['playlist'].apply(assign_session)
            df_cons = df_cons.dropna(subset=['Session_ID'])

            # 3. POOLING: Z-Score über die GESAMTE Session pro Proband berechnen
            # Wir gruppieren NUR nach Proband und Session (nicht nach Playlist!)
            pooled_data = []
            
            # Wir nutzen die Metrik, die oben im Tab "Verlauf" ausgewählt wurde (z.B. target_metric_9)
            # Falls die Variable anders heißt, hier anpassen. Ich nehme an, es ist die Metrik aus dem Lern-Effekt Block.
            calc_metric_cons = target_metric_9 

            for (pid, sess_id), group in df_cons.groupby(['participant', 'Session_ID']):
                raw_vals = group[calc_metric_cons].values
                
                # Z-Score Berechnung (Manuell oder via Helper)
                mean_sess = np.mean(raw_vals)
                std_sess = np.std(raw_vals)
                
                if std_sess == 0:
                    z_scores = np.zeros(len(raw_vals))
                else:
                    z_scores = (raw_vals - mean_sess) / std_sess
                
                temp = group.copy()
                temp['Session_Z_Score'] = z_scores
                pooled_data.append(temp)

            if pooled_data:
                df_pooled = pd.concat(pooled_data)

                # 4. PIVOTIEREN
                df_pivot = df_pooled.pivot_table(
                    index=['participant', 'item'], 
                    columns='Session_ID', 
                    values='Session_Z_Score',
                    aggfunc='mean'
                ).reset_index()

                df_pivot = df_pivot.dropna(subset=['Session 1', 'Session 2'])

                # ---------------------------------------------------------
                # GENRE & LABEL
                # ---------------------------------------------------------
                def resolve_genre(item_name):
                    if 'Grundaktivität' in str(item_name):
                        return 'Ruhephase'
                    return DEFAULT_GENRES.get(item_name, 'Unbekannt')

                df_pivot['genre'] = df_pivot['item'].apply(resolve_genre)
                df_pivot['Song_Label'] = df_pivot['item'] + " (" + df_pivot['genre'] + ")"

                # 5. METRIKEN BERECHNEN
                df_pivot['Delta_Abs'] = abs(df_pivot['Session 1'] - df_pivot['Session 2'])
                
                df_pivot['Is_Flip'] = (
                    ((df_pivot['Session 1'] > 0) & (df_pivot['Session 2'] < 0)) | 
                    ((df_pivot['Session 1'] < 0) & (df_pivot['Session 2'] > 0))
                )
                df_pivot['Consistency_Type'] = df_pivot['Is_Flip'].apply(lambda x: 'Instabil (Flip)' if x else 'Stabil')

                # ---------------------------------------------------------
                # VISUALISIERUNG
                # ---------------------------------------------------------
                
                st.markdown(f"#### Analyse für Metrik: **{calc_metric_cons}**")
                
                col_c1, col_c2 = st.columns([2, 1])

                with col_c1:
                    st.markdown("**A) Streuung: Reliabilität der Wirkung**")
                    st.caption("Vergleich der Z-Scores pro Song und Proband.")
                    
                    # Definiere die Sortierreihenfolge manuell für "Natural Sort"
                    # Wir gehen von Song 1 bis Song 12 aus (oder so viele wie da sind)
                    sort_order = [f"Song {i}" for i in range(1, 15)] 

                    chart_scatter = alt.Chart(df_pivot).mark_circle(size=70).encode(
                        x=alt.X('Session 1', title='Z-Score Session 1'),
                        y=alt.Y('Session 2', title='Z-Score Session 2'),
                        color=alt.Color('item', legend=alt.Legend(title="Song Name"), sort=sort_order), 
                        tooltip=['participant', 'item', 'genre', 'Session 1', 'Session 2', 'Delta_Abs']
                    ).interactive()

                    line = alt.Chart(pd.DataFrame({'x': [-3, 3], 'y': [-3, 3]})).mark_line(color='grey', strokeDash=[5,5]).encode(x='x', y='y')
                    st.altair_chart(chart_scatter + line, width='stretch')

                with col_c2:
                    st.markdown("**B) Stabilität (Richtung)**")
                    st.caption("Flip-Rate")
                    
                    flip_counts = df_pivot['Consistency_Type'].value_counts().reset_index()
                    flip_counts.columns = ['Typ', 'Anzahl']
                    
                    chart_pie = alt.Chart(flip_counts).mark_arc(innerRadius=50).encode(
                        theta=alt.Theta(field="Anzahl", type="quantitative"),
                        color=alt.Color(field="Typ", type="nominal", scale=alt.Scale(domain=['Stabil', 'Instabil (Flip)'], range=['#2ECC71', '#E74C3C'])),
                        tooltip=['Typ', 'Anzahl']
                    )
                    st.altair_chart(chart_pie, width='stretch')

                st.markdown("---")
                st.markdown("**C) Kontext-Ranking: Welche Songs lassen sich beeinflussen?**")
                
                # Berechnung des Durchschnitts
                song_context_stats = df_pivot.groupby('Song_Label')['Delta_Abs'].mean().reset_index()
                col_name_delta = 'Ø Delta (Kontext-Einfluss)'
                song_context_stats.columns = ['Song (Genre)', col_name_delta]
                
                song_context_stats = song_context_stats.set_index('Song (Genre)')
                
                # 1. Stabil: Kleiner als 1.0
                df_stable = song_context_stats[song_context_stats[col_name_delta] < 1.0].sort_values(col_name_delta, ascending=True)
                
                # 2. Instabil: Größer oder gleich 1.0
                df_instable = song_context_stats[song_context_stats[col_name_delta] >= 1.0].sort_values(col_name_delta, ascending=False)

                c_res, c_cham = st.columns(2)
                
                with c_res:
                    st.caption(f"**Stabil (Delta < 1.0)**: {len(df_stable)} Songs")
                    if not df_stable.empty:
                        st.dataframe(
                            df_stable.style.background_gradient(
                                subset=[col_name_delta], 
                                cmap='Greens_r',  # 0 (dunkelgrün) bis 1 (heller)
                                vmin=0, vmax=1.0  
                            )
                            .format({col_name_delta: '{:.3f}'})
                        )
                    else:
                        st.info("Keine Songs unter 1.0")

                with c_cham:
                    st.caption(f"**Kontext-Sensitiv (Delta ≥ 1.0)**: {len(df_instable)} Songs")
                    if not df_instable.empty:
                        st.dataframe(
                            df_instable.style.background_gradient(
                                subset=[col_name_delta], 
                                cmap='Reds',
                                vmin=1.0, vmax=max(2.0, df_instable[col_name_delta].max()) # Dynamische Obergrenze
                            )
                            .format({col_name_delta: '{:.3f}'})
                        )
                    else:
                        st.success("Keine instabilen Songs gefunden.")

# --------------------------
# DATEN VERARBEITUNG 
# --------------------------
def process_data(participants, cleaning_lvl, short_radius, react_buffer):
    master_data = []
    progress = st.progress(0)
    
    OPTIONAL_COLS = ['item_entspannt', 'item_bekannt', 'item_gefallen']
    
    # Baseline zuordnung
    PLAYLISTS_STD = ['audio1', 'audio2', 'youtube'] # normal (1 Termin)
    PLAYLISTS_SHFL = ['audio3', 'audio4', 'youtube_shuffle'] # shuffle (2 Termin)
    
    # Mapping: File-Key -> CSV ranking_type Name
    # Wir mappen 'youtube_shuffle' auf 'youtube', da in der CSV immer nur 'youtube' als ranking_type steht.
    R_MAP = {
        'audio1': 'audio_1', 
        'audio2': 'audio_2', 
        'audio3': 'audio_3', 
        'audio4': 'audio_4',
        'youtube': 'youtube',
        'youtube_shuffle': 'youtube' 
    }

    for idx, (p_id, files) in enumerate(participants.items()):
        if 'export' not in files: continue
        
        try:
            # 1. Ranking Datei laden
            ranking_df = pd.read_csv(files['export'])
            ranking_df.columns = ranking_df.columns.str.strip()
            
            # 2. Baselines vorbereiten
            ratios_std = None
            ratios_shfl = None
            
            # A) Standard Baseline (für Audio 1, 2, Youtube normal)
            if 'bl' in files:
                bl_df = load_eeg_file(files['bl'], "BL_STD")
                bl_ev = get_global_events(bl_df)
                bl_pwr = compute_band_power(clean_segment(bl_df, cleaning_lvl, short_radius, react_buffer, bl_ev), SAMPLING_FREQUENCY)
                if bl_pwr: ratios_std = calculate_ratios(bl_pwr)
            
            # B) Shuffle Baseline (für Audio 3, 4, Youtube Shuffle)
            if 'bl_shuffle' in files:
                bl_s_df = load_eeg_file(files['bl_shuffle'], "BL_SHFL")
                bl_s_ev = get_global_events(bl_s_df)
                bl_s_pwr = compute_band_power(clean_segment(bl_s_df, cleaning_lvl, short_radius, react_buffer, bl_s_ev), SAMPLING_FREQUENCY)
                if bl_s_pwr: ratios_shfl = calculate_ratios(bl_s_pwr)

            # 3. Playlists verarbeiten
            all_playlists = PLAYLISTS_STD + PLAYLISTS_SHFL
            
            for pl_key in all_playlists:
                if pl_key not in files: continue
                
                # ENTSCHEIDUNG: Baseline Wahl anhand der Gruppenzugehörigkeit
                is_shuffle_pl = pl_key in PLAYLISTS_SHFL
                current_ratios = ratios_shfl if is_shuffle_pl else ratios_std
                
                # Baseline Check
                if current_ratios is None: continue

                # Playlist EEG laden
                pl_df = load_eeg_file(files[pl_key], pl_key)
                pl_ev = get_global_events(pl_df)
                segments = segment_song_data(pl_df)
                
                # CSV Mapping
                csv_type = R_MAP.get(pl_key)
                subset = ranking_df[ranking_df['ranking_type'] == csv_type]
                if subset.empty: continue

                # --- HELPER ---
                def get_extras(item_name):
                    res = {}
                    for col in OPTIONAL_COLS:
                        if col in subset.columns:
                            val = subset[subset['item_name'] == item_name][col]
                            res[col] = val.iloc[0] if not val.empty else None
                        else:
                            res[col] = None 
                    return res
                # --------------

                # I. Grundaktivität
                try: ga_rank = subset[subset['item_name'] == 'Grundaktivität']['final_rank'].min()
                except: ga_rank = 6
                
                pl_items = []
                pl_items.append({
                    'participant': p_id, 
                    'playlist': pl_key, # Speichert z.B. 'youtube_shuffle' als Identifier
                    'item': 'Grundaktivität',
                    'is_baseline': True, 
                    'subj_rank': ga_rank, 
                    **current_ratios,
                    **get_extras('Grundaktivität')
                })
                
                # II. Songs
                song_subset = subset[~subset['item_name'].isin(['Grundaktivität', 'baseline'])].sort_values('play_order')
                song_names = song_subset['item_name'].tolist()
                
                for i, seg in enumerate(segments):
                    if i >= len(song_names): break
                    
                    s_pwr = compute_band_power(clean_segment(seg, cleaning_lvl, short_radius, react_buffer, pl_ev), SAMPLING_FREQUENCY)
                    if s_pwr:
                        s_ratios = calculate_ratios(s_pwr)
                        s_name = song_names[i]
                        s_rank = subset[subset['item_name'] == s_name]['final_rank'].min()
                        
                        pl_items.append({
                            'participant': p_id, 
                            'playlist': pl_key, 
                            'item': s_name,
                            'is_baseline': False, 
                            'subj_rank': s_rank, 
                            **s_ratios,
                            **get_extras(s_name)
                        })
                
                df_pl = pd.DataFrame(pl_items)
                for m in METRICS: 
                    df_pl[f'rank_eeg_{m}'] = df_pl[m].rank(ascending=False)
                
                master_data.append(df_pl)
                
        except Exception as e: 
            print(f"Fehler bei {p_id}: {e}")
            pass
            
        progress.progress((idx + 1) / len(participants))
    
    if master_data:
        st.session_state['master_data'] = pd.concat(master_data)
        st.success("Verarbeitung abgeschlossen! (Inkl. Youtube Shuffle)")
        st.rerun()
    else:
        st.error("Keine validen Daten generiert.")

if __name__ == "__main__":
    main()