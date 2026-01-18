import streamlit as st
import pandas as pd
import numpy as np
from scipy.signal import welch
import base64
from datetime import datetime

# --------------------------
# KONSTANTEN & SETUP
# --------------------------
SAMPLING_FREQUENCY = 250
EEG_CHANNELS = ["EEG 1", "EEG 2", "EEG 3", "EEG 4", "EEG 5", "EEG 6", "EEG 7", "EEG 8"]
CLEANING_LEVELS = ['Normal', 'Strikt', 'Sehr Strikt', 'Ohne']

# Pflichtspalten für Validierung
REQUIRED_EEG_COLS = EEG_CHANNELS + ['TRIG']
REQUIRED_RANKING_COLS = [
    'ranking_type', 'play_order', 'item_name', 'final_rank', 
    'participant_id', 'global_medikamente', 'global_operationen', 'global_musikeffekt'
]

EVENT_TRIGGERS = {
    'n': ('1', 'Song Start'), 'm': ('2', 'Song or Playlist End'),
    'q': ('3', 'Visuelles Event Start'), 'w': ('4', 'Visuelles Event Ende'),
    'a': ('5', 'Auditiv Event Start'), 's': ('6', 'Auditiv Event Ende'),
    'y': ('7', 'Bewegung Start'), 'x': ('8', 'Bewegung Ende'),
    'v': ('9', 'Transition Start'), 'b': ('10', 'Transition Ende'),
}

EVENT_NAME_MAP = {val[0]: val[1].replace(' Start', '').replace(' Ende', '') for val in {
    'n': ('1', 'Song'), 'm': ('2', 'Song'),
    'q': ('3', 'Visuelles Event'), 'w': ('4', 'Visuelles Event'),
    'a': ('5', 'Auditiv Event'), 's': ('6', 'Auditiv Event'),
    'y': ('7', 'Bewegung'), 'x': ('8', 'Bewegung'),
    'v': ('9', 'Transition'), 'b': ('10', 'Transition'),
}.values()}

# --------------------------
# HILFSFUNKTIONEN (BERECHNUNGEN)
# --------------------------

# --------------------------
# LOGIK-FUNKTIONEN 
# --------------------------

def calculate_dynamic_threshold(values, factor=0.5):
    """Berechnet einen relativen Toleranzwert (Standardabweichung * Faktor)."""
    if len(values) < 2: return 0.001
    return np.std(values) * factor

def get_strict_ranking_order(df, metric):
    """
    Erstellt eine Liste von Items basierend auf strikter numerischer Sortierung.
    Der Threshold wird ignoriert; der höchste Wert steht immer oben (Platz 1).
    """
    # Sortieren nach Wert absteigend
    df_sorted = df.sort_values(by=metric, ascending=False).copy()
    items = df_sorted['Phase'].values
    
    return items.tolist()

def get_sehr_fair_ranking_order(df, metric, user_ranking_list, threshold):
    """
    Sortiert die EEG-Items so um, dass sie dem User-Ranking möglichst nahe kommen,
    sofern die physiologische Differenz innerhalb der Toleranz liegt.
    """
    # Mapping User-Ranking: Item -> Rang (0-basiert)
    user_rank_map = {item: i for i, item in enumerate(user_ranking_list)}
    
    # Daten vorbereiten
    df_sorted = df.sort_values(by=metric, ascending=False).copy()
    values = df_sorted[metric].values
    items = df_sorted['Phase'].values
    
    # Optimierte Liste initialisieren (mit Platzhaltern)
    optimized_items = [None] * len(items)
    processed_indices = set()
    
    i = 0
    while i < len(values):
        # Cluster finden (Werte innerhalb Threshold)
        cluster_indices = [i]
        j = i + 1
        while j < len(values) and abs(values[i] - values[j]) < threshold:
            cluster_indices.append(j)
            j += 1
            
        # Items im Cluster
        cluster_item_names = [items[idx] for idx in cluster_indices]
        
        # Sortieren dieser Items nach ihrem Rang im User-Ranking (Best Match Logic)
        # Items, die nicht im User-Ranking sind, kommen ans Ende des Clusters
        cluster_item_names.sort(key=lambda x: user_rank_map.get(x, 999))
        
        # In die Ergebnisliste schreiben
        for k, item_name in enumerate(cluster_item_names):
            # Der Slot im Gesamtranking entspricht dem Startindex des Clusters + k
            optimized_items[i + k] = item_name
            
        i = j # Sprung zum nächsten Cluster
        
    return optimized_items

def validate_columns(df, required_cols, file_name):
    """Prüft, ob alle notwendigen Spalten im DataFrame vorhanden sind."""
    df.columns = df.columns.str.strip()
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        return f"❌ Datei '{file_name}': Fehlende Spalten: {', '.join(missing)}"
    return None

def determine_dominant_metric(scores):
    ab_score = scores.get('Alpha/Beta Verhältnis', 0)
    tb_score = scores.get('Theta/Beta Verhältnis', 0)
    atb_score = scores.get('(Alpha+Theta)/Beta Verhältnis', 0)
    if ab_score == tb_score and ab_score >= atb_score and ab_score > 0: return '(Alpha+Theta)/Beta Verhältnis'
    if ab_score == atb_score and ab_score > tb_score: return 'Alpha/Beta Verhältnis'
    if tb_score == atb_score and tb_score > ab_score: return '(Alpha+Theta)/Beta Verhältnis'
    if not scores or all(v == 0 for v in scores.values()): return None
    return max(scores, key=scores.get)

def compute_band_power(data_segment, sf, eeg_channels):
    bands = {'Theta': (4, 8), 'Alpha': (8, 13), 'Beta': (13, 30)}
    eeg_data = data_segment[eeg_channels].values.T
    if eeg_data.shape[1] < sf: return pd.DataFrame(index=eeg_channels, columns=bands.keys()).fillna(0)
    freqs, psd = welch(eeg_data, sf, nperseg=sf)
    return pd.DataFrame({band: np.mean(psd[:, np.logical_and(freqs >= low, freqs <= high)], axis=1) for band, (low, high) in bands.items()}, index=eeg_channels)

def calculate_relaxation_ratios(band_power_df):
    if band_power_df.empty or any(b not in band_power_df.columns for b in ['Alpha', 'Beta', 'Theta']): return {'Alpha/Beta Verhältnis': 0.0, 'Theta/Beta Verhältnis': 0.0, '(Alpha+Theta)/Beta Verhältnis': 0.0}
    totals = {band: band_power_df[band].sum() for band in ['Alpha', 'Theta', 'Beta']}
    if totals['Beta'] == 0: return {'Alpha/Beta Verhältnis': np.inf, 'Theta/Beta Verhältnis': np.inf, '(Alpha+Theta)/Beta Verhältnis': np.inf}
    return {'Alpha/Beta Verhältnis': totals['Alpha'] / totals['Beta'], 'Theta/Beta Verhältnis': totals['Theta'] / totals['Beta'], '(Alpha+Theta)/Beta Verhältnis': (totals['Alpha'] + totals['Theta']) / totals['Beta']}

def analyze_playlist_triggers(df):
    trigger_col = 'TRIG'; all_triggers = df[df[trigger_col] != 0]
    starts = all_triggers[all_triggers[trigger_col] == 1].index.tolist()
    ends = all_triggers[all_triggers[trigger_col] == 2].index.tolist()
    if not starts: return [], "Keine Start-Trigger (TRG=1) gefunden."
    master_end_point = len(df); status_message = "Playlist Ende = Dateiende"
    if ends and ends[-1] > starts[-1]:
        master_end_point = ends[-1]; status_message = "Playlist End-Trigger gefunden"
    song_trigger_indices = sorted([idx for idx in starts + ends if idx < master_end_point])
    segments = []
    for i, start_index in enumerate(starts):
        if start_index >= master_end_point: continue
        segment_end = master_end_point
        for trigger_index in song_trigger_indices:
            if trigger_index > start_index: segment_end = trigger_index; break
        segments.append(df.iloc[start_index:segment_end])
    return segments, status_message

def get_global_events(playlist_df):
    triggers_in_df = playlist_df[playlist_df['TRIG'] != 0]['TRIG']
    all_pairs = {'3': '4', '5': '6', '7': '8', '9': '10'}
    long_events, transitions, short_events = [], [], []
    processed_indices = set()
    trigger_indices = triggers_in_df.index.tolist()
    trigger_values = triggers_in_df.astype(int).astype(str).tolist()

    try:
        if '1' in trigger_values:
            first_song_start_index_pos = trigger_values.index('1')
            first_song_start_df_idx = trigger_indices[first_song_start_index_pos]
            next_limit_pos = len(trigger_values)
            if '1' in trigger_values[first_song_start_index_pos + 1:]:
                next_limit_pos = trigger_values.index('1', first_song_start_index_pos + 1)
            for i in range(first_song_start_index_pos + 1, next_limit_pos):
                val = trigger_values[i]
                if val == '9': break 
                if val == '10':
                    current_df_idx = trigger_indices[i]
                    event_pair = (first_song_start_df_idx, current_df_idx, "Anfangs-Transition")
                    transitions.append(event_pair)
                    processed_indices.add(current_df_idx) 
                    break
    except ValueError: pass

    for i, start_val in enumerate(trigger_values):
        start_idx = trigger_indices[i]
        if start_idx in processed_indices: continue
        if start_val in all_pairs:
            end_val = all_pairs[start_val]
            event_type = EVENT_NAME_MAP.get(start_val, "Unbekannt")
            for j in range(i + 1, len(trigger_values)):
                end_idx = trigger_indices[j]
                if end_idx in processed_indices: continue
                if trigger_values[j] == end_val:
                    event_pair = (start_idx, end_idx, event_type)
                    if start_val == '9': transitions.append(event_pair)
                    else: long_events.append(event_pair)
                    processed_indices.add(start_idx); processed_indices.add(end_idx)
                    break
                    
    for i, val in enumerate(trigger_values):
        idx = trigger_indices[i]
        event_type = EVENT_NAME_MAP.get(val, "Unbekannt")
        if idx not in processed_indices and val not in ['1', '2']:
            short_events.append((idx, event_type))
            
    return long_events, transitions, short_events

def clean_segment_with_context(segment, cleaning_level, short_event_seconds, reaction_buffer, global_long, global_trans, global_short):
    if cleaning_level == 'Ohne' or segment.empty: return segment
    indices_to_drop = set()
    seg_start, seg_end = segment.index.min(), segment.index.max()
    
    # Umrechnung Sekunden in Samples (Zeilen)
    buffer_rows = int(reaction_buffer * SAMPLING_FREQUENCY)
    
    # LEVEL: NORMAL (Basis-Bereinigung: Nur Transitions entfernen)
    # Transitions bleiben strikt (Start bis Ende), da diese keine "unerwarteten" Events sind
    if cleaning_level in ['Normal', 'Strikt', 'Sehr Strikt']:
        for start, end, _ in global_trans:
            if start <= seg_end and end >= seg_start:
                drop_start = max(start, seg_start)
                drop_end = min(end, seg_end)
                indices_to_drop.update(segment.loc[drop_start:drop_end].index)
    
    # LEVEL: STRIKT (Normal + Lange Events)
    # Hier wird der Startpunkt um den Reaktions-Puffer nach vorne verlegt
    if cleaning_level in ['Strikt', 'Sehr Strikt']:
        for start, end, _ in global_long:
            if start <= seg_end and end >= seg_start:
                # Puffer anwenden: Startpunkt liegt früher (links auf Zeitachse)
                adjusted_start = start - buffer_rows
                
                drop_start = max(adjusted_start, seg_start)
                drop_end = min(end, seg_end)
                indices_to_drop.update(segment.loc[drop_start:drop_end].index)

    # LEVEL: SEHR STRIKT (Strikt + Kurze Events)
    # Hier wird der Puffer additiv auf den Radius für die linke Seite angewendet
    if cleaning_level == 'Sehr Strikt':
        rows_to_cut = int(short_event_seconds * SAMPLING_FREQUENCY)
        for idx, _ in global_short:
            if seg_start <= idx <= seg_end:
                # Links: Radius + Puffer | Rechts: Nur Radius
                drop_start = max(seg_start, idx - (rows_to_cut + buffer_rows))
                drop_end = min(seg_end, idx + rows_to_cut)
                indices_to_drop.update(segment.loc[drop_start:drop_end].index)
                
    return segment.drop(index=list(indices_to_drop), errors='ignore')

def count_events(df):
    trigger_col = 'TRIG'
    if trigger_col not in df.columns: return {}
    
    long_events, transitions, short_events = get_global_events(df)
    trigger_values = df[df[trigger_col] != 0][trigger_col].astype(int).astype(str).tolist()
    
    # 1. Zähle rein physikalische Trigger für die Anzeige
    raw_9_count = trigger_values.count('9')
    raw_10_count = trigger_values.count('10')
    
    # 2. Ermittle, wie viele "Sonderfälle" (1 -> 10) der Algorithmus gefunden hat
    special_case_count = sum(1 for _, _, name in transitions if name == "Anfangs-Transition")

    counts = {
        'Song Start': trigger_values.count('1'), 
        'Song or Playlist End': trigger_values.count('2'),
        'Visuelles Event': {'Lange': 0, 'Kurze': 0}, 
        'Auditiv Event': {'Lange': 0, 'Kurze': 0},
        'Bewegung': {'Lange': 0, 'Kurze': 0}, 
        'Transition': {'Start': raw_9_count, 'Ende': raw_10_count, 'Special': special_case_count}
    }
    for _, _, event_type in long_events:
        if event_type in counts: counts[event_type]['Lange'] += 1
    for _, event_type in short_events:
        if event_type == 'Transition': continue
        if event_type in counts and 'Kurze' in counts[event_type]: counts[event_type]['Kurze'] += 1
    return counts

def display_event_counts(counts, file_name="default", end_status=None, is_baseline=False):
    if not counts: return
    has_events = any(val > 0 for k, val in counts.items() if isinstance(val, int)) or \
                 any(d.get('Lange', 0) > 0 or d.get('Kurze', 0) > 0 for k, d in counts.items() if k != 'Transition' and isinstance(d, dict)) or \
                 (counts.get('Transition', {}).get('Start', 0) > 0 or counts.get('Transition', {}).get('Ende', 0) > 0)

    if not has_events: st.info("Keine Events in dieser Datei gefunden."); return
    start_label = "Baseline Start" if is_baseline else "Song Start"
    end_label = "Baseline Ende" if is_baseline else "Song Ende"

    st.markdown(f"**{start_label}:** `{counts.get('Song Start', 0)}`")
    st.markdown(f"**{end_label}:** `{counts.get('Song or Playlist End', 0)}`")
    
    vis_counts = counts.get('Visuelles Event', {}); st.markdown(f"**Visuelle Events:** Kurz: `{vis_counts.get('Kurze',0)}` / Lang: `{vis_counts.get('Lange',0)}`")
    aud_counts = counts.get('Auditiv Event', {}); st.markdown(f"**Auditive Events:** Kurz: `{aud_counts.get('Kurze',0)}` / Lang: `{aud_counts.get('Lange',0)}`")
    mov_counts = counts.get('Bewegung', {}); st.markdown(f"**Bewegungs-Events:** Kurz: `{mov_counts.get('Kurze',0)}` / Lang: `{mov_counts.get('Lange',0)}`")
    
    trans_counts = counts.get('Transition', {'Start': 0, 'Ende': 0, 'Special': 0})
    t_start = trans_counts.get('Start', 0)
    t_end = trans_counts.get('Ende', 0)
    t_special = trans_counts.get('Special', 0)

    if is_baseline:
        if t_start > 0 or t_end > 0: st.error(f"⚠️ Warnung: Transition-Trigger in Baseline gefunden! (Start: {t_start}, Ende: {t_end})")
    else:
        st.markdown(f"**Transition Start:** `{t_start}`")
        st.markdown(f"**Transition Ende:** `{t_end}`")
        
        # Validierungslogik
        if t_start == t_end:
            if t_special > 0:
                # Fall: 1->10 und 9->10 existieren beide sauber, aber wir haben mehr Ends als 9er starts? 
                # Eigentlich: (Start + Special) sollte == Ende sein.
                pass 
        elif t_start < t_end:
            missing_starts = t_end - t_start
            if missing_starts == t_special:
                st.info(f"ℹ️ Info: {t_special}x Anfangs-Transition erkannt (Trigger 1 ➔ 10). Gleicht die fehlenden Start-Trigger aus.")
            else:
                st.error(f"⚠️ Warnung: Mehr Transition-Enden ({t_end}) als Starts ({t_start}). (Special Cases: {t_special})")
        else: # Start > End
             st.error(f"⚠️ Warnung: Mehr Transition-Starts ({t_start}) als Enden ({t_end})!")

    if end_status: 
        if "Invalide" in end_status: st.error(end_status)
        elif "Valide" in end_status: st.success(end_status)
        else: st.caption(end_status)

def get_rankings_from_df(df, ranking_type):
    if df.empty or 'ranking_type' not in df.columns: 
        return [], []
    
    ranking_df = df[df['ranking_type'] == ranking_type].copy()
    
    preference_df = ranking_df.sort_values(by='final_rank')
    preference_list = preference_df['item_name'].tolist()

    mask_valid_songs = pd.Series(True, index=ranking_df.index)
    
    if 'item_id' in ranking_df.columns:
        mask_valid_songs &= (ranking_df['item_id'] != 'baseline')
    
    mask_valid_songs &= (ranking_df['item_name'] != 'Grundaktivität')
    
    chrono_df = ranking_df[mask_valid_songs].copy()
    
    # Nach play_order sortieren
    chrono_df = chrono_df.sort_values(by='play_order')
    
    chronological_list = chrono_df['item_name'].tolist()
    
    return chronological_list, preference_list

# Hilfsfunktion für Dateipersistenz
def persist_uploaded_file(upload_obj, key):
    if upload_obj is not None:
        st.session_state[key] = upload_obj
    return st.session_state.get(key)

# --------------------------
# REPORT GENERATOR 
# --------------------------

def generate_html_report(results_collection, short_event_seconds, reaction_buffer, participant_id):
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")
    
    # Berechne Gesamtzeit links für die Info
    total_left_cut = short_event_seconds + reaction_buffer
    
    logic_desc = f"""
    <div style="background: #eef; padding: 10px; border-left: 4px solid #44a; margin-bottom: 20px; font-size: 0.9em;">
        <strong>Angewandte Reinigungs-Logik:</strong><br>
        <ul>
            <li><b>Ohne:</b> Keine Reinigung und nur Segmentierung der Songs.</li>
            <li><b>Normal:</b> Entfernt Transitionen.</li>
            <li><b>Strikt:</b> Entfernt Transitionen + Lange Störungen (Startpunkt um {reaction_buffer}s vorverlegt).</li>
            <li><b>Sehr Strikt:</b> Entfernt Transitionen + Lange Störungen + Kurze Störungen.<br>
            <i>(Schnitt bei Kurz-Events: {total_left_cut}s Vergangenheit [Radius {short_event_seconds}s + Puffer {reaction_buffer}s] bis {short_event_seconds}s Zukunft)</i>.</li>
        </ul>
    </div>
    """

    html = f"""
    <html>
    <head>
        <title>EEG Analyse Report</title>
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; margin: 20px; }}
            h1 {{ color: #2c3e50; border-bottom: 2px solid #eee; padding-bottom: 10px; }}
            h2 {{ color: #16a085; margin-top: 30px; border-bottom: 1px solid #ddd; padding-bottom: 5px; }}
            h3 {{ color: #2980b9; margin-top: 20px; font-size: 1.1em; }}
            .level-section {{ margin-bottom: 60px; border: 1px solid #e0e0e0; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
            .level-header {{ background-color: #2c3e50; color: white; padding: 10px 20px; border-radius: 5px 5px 0 0; margin: -20px -20px 20px -20px; font-size: 1.4em; }}
            .meta-info {{ background: #f8f9fa; padding: 15px; border-radius: 5px; margin-bottom: 20px; border: 1px solid #ddd; }}
            table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 0.9em; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            th {{ background-color: #f2f2f2; color: #333; }}
            tr:nth-child(even) {{ background-color: #f9f9f9; }}
            .match-green {{ background-color: #28a745; color: white; font-weight: bold; }}
            .match-red {{ background-color: #dc3545; color: white; }}
            .summary-box {{ background: #e8f6f3; padding: 15px; border-left: 5px solid #16a085; margin-bottom: 20px; }}
            .metric-box {{ display: inline-block; background: white; padding: 10px; margin-right: 10px; border: 1px solid #ddd; border-radius: 4px; }}
            .footer {{ margin-top: 50px; color: #888; font-size: 0.8em; text-align: center; border-top: 1px solid #eee; padding-top: 10px; }}
        </style>
    </head>
    <body>
        <h1>🧠 EEG Analyse Report</h1>
        <div class="meta-info">
            <strong>Teilnehmer ID:</strong> {participant_id}<br>
            <strong>Erstellt am:</strong> {now_str}<br>
            <strong>Enthaltene Level:</strong> {', '.join(results_collection.keys())}<br>
            <strong>Kurze Events-Radius:</strong> {short_event_seconds} Sek.<br>
            <strong>Reaktions-Puffer (Rückwirkend):</strong> {reaction_buffer} Sek.
        </div>
        {logic_desc}
    """
    for cleaning_level, all_playlist_results in results_collection.items():
        html += f"""
        <div class="level-section">
        <div class="level-header">Reinigungs-Level: {cleaning_level}</div>
        """
        total_match_ranks = 0; total_items_checked = 0
        individual_matches = {'Alpha/Beta Verhältnis': 0, 'Theta/Beta Verhältnis': 0, '(Alpha+Theta)/Beta Verhältnis': 0}
        theoretical_max = 0; valid_analysis_exists = False
        
        for r_type, data in all_playlist_results.items():
            results_df = data['eeg_results_df']; pref_ranking = data['preference_ranking']
            eeg_items = set(results_df['Phase']); common_items = list(eeg_items.intersection(set(pref_ranking)))
            if not common_items: continue
            valid_analysis_exists = True
            
            filtered_pref = [i for i in pref_ranking if i in common_items]
            filtered_eeg = results_df[results_df['Phase'].isin(common_items)]
            num_items = len(filtered_pref); total_items_checked += num_items
            
            metrics = list(individual_matches.keys())
            eeg_rankings = {m: filtered_eeg.sort_values(by=m, ascending=False)['Phase'].tolist() for m in metrics}
            current_theoretical_max = num_items
            for i in range(num_items):
                user_item = filtered_pref[i]
                rank_has_match = False
                eeg_suggestions = []
                for m in metrics:
                    if i < len(eeg_rankings[m]):
                        eeg_item = eeg_rankings[m][i]; eeg_suggestions.append(eeg_item)
                        if user_item == eeg_item: individual_matches[m] += 1; rank_has_match = True
                if rank_has_match: total_match_ranks += 1
                bonus = len(eeg_suggestions) - len(set(eeg_suggestions)); current_theoretical_max += bonus
            theoretical_max += current_theoretical_max

        if valid_analysis_exists:
            dominant = determine_dominant_metric(individual_matches)
            interpretations = {
                'Alpha/Beta Verhältnis': "Du assoziierst Entspannung am ehesten mit einem Zustand <b>wacher Aufmerksamkeit</b>.", 
                'Theta/Beta Verhältnis': "Du assoziierst Entspannung am ehesten mit einem <b>tiefen, meditativen Zustand</b>, der an Schläfrigkeit grenzt.", 
                '(Alpha+Theta)/Beta Verhältnis': "Du assoziierst Entspannung am ehesten mit einer <b>ausgeglichenen Mischung</b> aus wacher und tiefer Ruhe."
            }
            interp_text = interpretations.get(dominant, "Keine klare Tendenz.") if dominant else "Keine Daten."
            html += f"""
            <div class="summary-box">
                <h3>Gesamt-Interpretation ({cleaning_level})</h3>
                <p>{interp_text}</p>
                <div style="margin-top: 10px;">
                    <span class="metric-box"><b>Trefferquote:</b> {total_match_ranks}/{total_items_checked}</span>
                    <span class="metric-box"><b>Score:</b> {sum(individual_matches.values())}/{theoretical_max}</span>
                </div>
                <p><small>Detaillierte Treffer: {', '.join([f"{k.replace(' Verhältnis','')}: {v}" for k,v in individual_matches.items()])}</small></p>
            </div>
            """
        else: html += "<div class='summary-box'><p>Keine übereinstimmenden Daten für eine Gesamtanalyse gefunden.</p></div>"

        metrics = ['Alpha/Beta Verhältnis', 'Theta/Beta Verhältnis', '(Alpha+Theta)/Beta Verhältnis']
        for r_type, data in all_playlist_results.items():
            html += f"<h3>Playlist: {r_type.replace('_', ' ').title()}</h3>"
            results_df = data['eeg_results_df']; preference_ranking = data['preference_ranking']
            if results_df.empty: html += "<p>Keine EEG Ergebnisse.</p>"; continue
            html += "<h4>📊 EEG Messwerte (Top-List)</h4>"
            html += "<table><thead><tr><th>Rang</th>"
            for m in metrics: html += f"<th>{m}</th>"
            html += "</tr></thead><tbody>"
            
            sorted_lists = {m: results_df.sort_values(by=m, ascending=False) for m in metrics}
            max_len = len(results_df)
            for i in range(max_len):
                html += f"<tr><td>{i+1}</td>"
                for m in metrics:
                    try: row = sorted_lists[m].iloc[i]; html += f"<td><b>{row['Phase']}</b> <small>[{row[m]:.3f}]</small></td>"
                    except: html += "<td>-</td>"
                html += "</tr>"
            html += "</tbody></table>"
            
            html += "<h4>🎯 Vergleich mit Subjektivem Ranking</h4>"
            eeg_items_set = set(results_df['Phase'])
            common_items = [item for item in preference_ranking if item in eeg_items_set]
            
            if common_items:
                html += "<table><thead><tr><th>Rang</th><th>Dein Ranking</th>"
                for m in metrics: html += f"<th>{m.replace(' Verhältnis', '').replace('Alpha', 'A').replace('Theta', 'T').replace('Beta', 'B')}</th>"
                html += "</tr></thead><tbody>"
                eeg_rank_lists = {m: results_df.sort_values(by=m, ascending=False)['Phase'].tolist() for m in metrics}
                eeg_rank_filtered = {m: [x for x in lst if x in common_items] for m, lst in eeg_rank_lists.items()}
                for i, user_item in enumerate(common_items):
                    html += f"<tr><td>{i+1}</td><td><strong>{user_item}</strong></td>"
                    for m in metrics:
                        eeg_item = eeg_rank_filtered[m][i] if i < len(eeg_rank_filtered[m]) else "-"
                        css_class = "match-green" if eeg_item == user_item else "match-red"
                        html += f"<td class='{css_class}'>{eeg_item}</td>"
                    html += "</tr>"
                html += "</tbody></table>"
            else: html += "<p>Keine Übereinstimmung zwischen Playlist-Items und Ranking-Namen.</p>"
        html += "</div>"

    html += """<div class="footer">Generiert mit Streamlit EEG Analyzer</div></body></html>"""
    return html

# --------------------------
# HAUPTANWENDUNG
# --------------------------

def display_auswertung_page():
    # Initialisierung der State-Variablen für Dateipersistenz
    if 'stored_baseline' not in st.session_state: st.session_state.stored_baseline = None
    if 'stored_youtube' not in st.session_state: st.session_state.stored_youtube = None
    if 'stored_audio1' not in st.session_state: st.session_state.stored_audio1 = None
    if 'stored_audio2' not in st.session_state: st.session_state.stored_audio2 = None
    if 'stored_audio3' not in st.session_state: st.session_state.stored_audio3 = None 
    if 'stored_audio4' not in st.session_state: st.session_state.stored_audio4 = None 
    if 'stored_ranking_csv' not in st.session_state: st.session_state.stored_ranking_csv = None
    
    if 'analysis_results' not in st.session_state:
        st.session_state.analysis_results = None

    st.title("🧠 Analyse"); st.markdown("---") 
    
    # --------------------------
    # INFOBEREICHE 
    # --------------------------
    with st.expander("Grundverständnis zur Auswertung", expanded=False):
        st.markdown("- **Alpha/Beta Verhältnis:** Indikator für **wache Entspannung**.\n- **Theta/Beta Verhältnis:** Indikator für **tiefere Entspannung/Schläfrigkeit**.\n- **(Alpha+Theta)/Beta Verhältnis:** Ein **kombinierter Index**.\n\n**Generell gilt:** Je höher das Verhältnis, desto ausgeprägter der relative Entspannungszustand.")
    
    with st.expander("Song Segmentierung ", expanded=False):
        st.markdown("""
        **1. Globale Grenzen & Trigger-Scan**
        Der Algorithmus identifiziert alle Song Start- (1) und Song End-Trigger (2). Der letzte End-Trigger in der Datei (falls vorhanden) definiert den *Master-Endpunkt*, nach dem keine neuen Segmente mehr gestartet werden (Ende der Playlist Aufnahme ➔ Rest der Datei ist irrelevant für die Analyse).

        **2. Dynamische Segment-Definition**
        Es wird durch jeden Start-Trigger iteriert. Das Ende eines Segments wird durch den *nächstgelegenen* Song-Trigger bestimmt – unabhängig davon, ob es sich um einen Start oder ein Ende handelt. *(Aufnahme einer Playlist ohne Pausen: [1 ➔ 1 ➔ 1 ➔ 1 ➔ 1 ➔ 2] )*

        **3. Abgedeckte Edge Cases**
        - **Pausen (1 ➔ 2 .... 1 ➔ 1):** Ein Song startet und wird regulär beendet, woraufhin nach einer Pause der nächste Song startet markiert wird.
        - **Fehlendes Ende (1 ➔ Datei-Ende):** Ein Song hat keinen End-Trigger. Das Song-Segment gilt bis zum Ende der Aufnahme/Datei.
        """)

    with st.expander("Events ", expanded=False):
        st.markdown("""
        Die Events (Start/Ende) bedeuten folgendes:
        - **Visuell (3/4):** Visuelle Ablenkungen (Lichtverhältnisse).
        - **Auditiv (5/6):** Störungen durch Geräusche (Husten, Sprechen, Lärm).
        - **Bewegung (7/8):** Störungen durch Kopf- oder Körperbewegung.
        - **Transition (9/10):** Die Übergangsphase zwischen zwei Songs, wodurch nach Reinigung nur noch der "Hauptteil" jedes Songs zur Analyse genutzt mit reduzierter Beeinflussung der Songs untereinander (z.B. 15 Sekunden vor Ende und nach Beginn).
                    
        ---
        
        Event-Trigger Logik:      
                      
        **1. Lange Events (Zeitspannen)**
        Der Algorithmus sucht nach **zusammengehörigen Paaren** (Start ➔ Ende).
        *   Findet er z.B. eine **3** und später eine **4**, wird der gesamte Bereich dazwischen als *langes visuelles Event* markiert.

        **2. Kurze Events (Punktuelle Störungen)**
        Trigger, die **keinen passenden Partner** haben (z.B. eine einzelne **5** ohne nachfolgende **6**), werden als *kurze Events* klassifiziert.
        *   Dies repräsentiert kurze Spikes (z.B. einmaliges Husten).
        *   Das heißt es können mehrere kurze Events hintereinander existieren [1 ➔ 5 ➔ 5 ➔ 1 ➔ 1 ➔ 5 ➔ 1 ➔ 1 ➔ 2]
        *   Bei überschneidenden Bereiche (z.B. [7 ➔ 7] mit einem Trigger Zeitabstand von 4 Sekunden und bei der Einstellung für kurze Events von 5 Sekunden ) "verschmelzen" einfach zu einem einzigen, längeren ausgeschnittenen Block.

        **3. Sonderfälle: Transition Trigger [1 ➔ 10 ➔ 9 ➔ 1 ➔ 10....]**
        *   **Anfangs-Transition:** Startet die Playlist (Trigger **1**) und es folgt darauf ein Transition-Ende (**10**), ohne dass vorher eine **9** kam gilt der Bereich **[1 ➔ 10]** als Übergangsphase und wird im "Sehr Strikt"-Reinigungs-Level entfernt.
        *   **Transition über mehrere Songs [1 ➔ 10 ➔ 9 ➔ 1 ➔ 1 ➔ 10....]:** Beim Reinigungs-Level würden in dem Fall ganze Songs rausgeschnitten werden. Wird jedoch als Fehlermeldung angezeigt wenn dieser Fall erkannt wird.
        """)

    with st.expander("Erklärung der Reinigungs-Level", expanded=False):
        st.markdown("""
        Die Analyse wird in vier Stufen durchgeführt, um Störungen unterschiedlich stark zu filtern:
        1. **Ohne:** Es werden die kompletten Rohdaten zwischen Start (1) und Ende (2) genutzt [Alternativ auch: Start und Start 1 ➔ 1]. Keine Filterung.
        2. **Normal:** Es werden **Transitionsphasen/Übergangsphasen** herausgeschnitten. (Am Anfang und Ende eines Songs, wenn noch Unruhe herrscht oder sich sonst die Lieder am stärksten gegenseitig beeinflussen.)
        3. **Strikt:** Bereiche, die als **"Lange Events"** (z.B. Bewegungen oder Lärm über einen längeren Zeitraum) markiert sind, werden zusätzlich herausgeschnitten.
        4. **Sehr Strikt:** Zusätzlich zu Strikt werden auch **"Kurze Events"** herausgeschnitten. Im angegebenen Sekunden Radius in beide Richtungen.
        """)
    
    st.markdown("---")
    st.subheader("1. Datenquelle für das subjektive Ranking auswählen")
    source_choice = st.radio("Datenquelle", ("Aktuelle Session-Daten verwenden", "CSV-Exportdatei hochladen"), horizontal=True, label_visibility="collapsed")
    
    user_ranking_df = pd.DataFrame()
    validation_errors = []
    
    # Global Infos Variables
    participant_id = "Unbekannt"
    global_meds = "N/A"
    global_ops = "N/A"
    global_music_fx = "N/A"

    # --- DATENQUELLE: SESSION DATEN ---
    if source_choice == "Aktuelle Session-Daten verwenden":
        if 'participant_id' in st.session_state:
            session_rows = []
            p_id = st.session_state.get('participant_id', 'N/A').strip()
            g_med = st.session_state.get('global_medikamente', 'N/A').strip()
            g_ops = st.session_state.get('global_operationen', 'N/A').strip()
            g_mus = 'Ja' if st.session_state.get('global_musikeffekt', False) else 'Nein'
            
            participant_id = p_id; global_meds = g_med; global_ops = g_ops; global_music_fx = g_mus
            
            # Ranking Map erweitert um Audio 3 & 4
            ranking_map = {
                'youtube': 'ranked_videos', 
                'audio_1': 'ranked_audio_1', 
                'audio_2': 'ranked_audio_2',
                'audio_3': 'ranked_audio_3',
                'audio_4': 'ranked_audio_4'
            }
            for r_type, state_key in ranking_map.items():
                items = st.session_state.get(state_key, [])
                if items:
                    for i, item in enumerate(items):
                        row = {
                            'ranking_type': r_type, 'play_order': item.get('play_order', 999),
                            'item_name': item.get('title') or item.get('name', 'N/A'), 'final_rank': i + 1,
                            'participant_id': p_id, 'global_medikamente': g_med,
                            'global_operationen': g_ops, 'global_musikeffekt': g_mus
                        }
                        session_rows.append(row)
            if session_rows: user_ranking_df = pd.DataFrame(session_rows)
            else: validation_errors.append("Keine Ranking-Daten in der aktuellen Session gefunden.")
        else: validation_errors.append("Keine aktive Session gefunden. Bitte zuerst Rankings im ersten Tab durchführen.")

    # --- DATENQUELLE: CSV UPLOAD ---
    elif source_choice == "CSV-Exportdatei hochladen":
        # Upload Logik mit Persistenz
        uploaded_csv_raw = st.file_uploader("Subjektive Ranking-CSV-Datei hochladen", type="csv")
        final_ranking_file = persist_uploaded_file(uploaded_csv_raw, 'stored_ranking_csv')
        
        if final_ranking_file:
            if uploaded_csv_raw is None: st.success(f"✅ Verwende gespeicherte Ranking-Datei: {final_ranking_file.name}")
            try: 
                final_ranking_file.seek(0) # Reset pointer
                user_ranking_df = pd.read_csv(final_ranking_file)
                user_ranking_df.columns = user_ranking_df.columns.str.strip()
                err = validate_columns(user_ranking_df, REQUIRED_RANKING_COLS, "Ranking Datei")
                if err: validation_errors.append(err)
                else:
                    if not user_ranking_df.empty:
                        participant_id = user_ranking_df['participant_id'].iloc[0]
                        global_meds = user_ranking_df['global_medikamente'].iloc[0]
                        global_ops = user_ranking_df['global_operationen'].iloc[0]
                        global_music_fx = user_ranking_df['global_musikeffekt'].iloc[0]
            except Exception as e: validation_errors.append(f"Fehler beim Lesen der Ranking-Datei: {e}")
    
    # --- TEILNEHMER DETAILS ANZEIGE ---
    if not user_ranking_df.empty and participant_id != "Unbekannt":
        with st.expander("👤 Teilnehmer Informationen", expanded=True):
            info_cols = st.columns(2)
            info_cols[0].markdown(f"**Teilnehmer ID:** `{participant_id}`")
            info_cols[0].markdown(f"**Medikamente:** `{global_meds}`")
            info_cols[1].markdown(f"**Operationen:** `{global_ops}`")
            info_cols[1].markdown(f"**Musikeffekt:** `{global_music_fx}`")

    # --------------------------
    # RANKING CSV MERGER TOOL 
    # --------------------------
    with st.expander("🛠️ Ranking CSV Merger Tool (Dateien zusammenfügen)", expanded=False):
        st.info("Hier kannst du mehrere exportierte Ranking-CSVs zu einer einzigen Datei zusammenfügen.")
        merge_files = st.file_uploader("CSV-Dateien zum Zusammenfügen auswählen", type="csv", accept_multiple_files=True, key="merge_uploader")
        
        if merge_files:
            merge_dfs = []
            merge_filenames = [] # Zum Zuordnen von Fehlern/Hinweisen
            parse_error = False
            
            # Dictionary zum Speichern, welcher Typ in welcher Datei vorkommt
            ranking_type_sources = {}
            
            # 1. Erfassen der Dateien
            for f in merge_files:
                try:
                    f.seek(0)
                    tmp_df = pd.read_csv(f)
                    tmp_df.columns = tmp_df.columns.str.strip()

                    if 'ranking_type' not in tmp_df.columns:
                        st.error(f"❌ Datei '{f.name}' ist ungültig (keine 'ranking_type' Spalte).")
                        parse_error = True
                    else:
                        merge_dfs.append(tmp_df)
                        merge_filenames.append(f.name)
                        
                        types_in_file = tmp_df['ranking_type'].dropna().unique()
                        for rt in types_in_file:
                            if rt not in ranking_type_sources:
                                ranking_type_sources[rt] = []
                            ranking_type_sources[rt].append(f.name)
                            
                except Exception as e:
                    st.error(f"❌ Fehler beim Lesen von '{f.name}': {e}")
                    parse_error = True
            
            if not parse_error and merge_dfs:
                
                # --- SPALTEN HARMONISIERUNG ---
                all_columns = set()
                for df in merge_dfs:
                    all_columns.update(df.columns)
                
                filled_columns_report = []
                
                for i, df in enumerate(merge_dfs):
                    missing_cols = list(all_columns - set(df.columns))
                    if missing_cols:
                        # Fehlende Spalten mit "N/A" auffüllen
                        for col in missing_cols:
                            df[col] = "N/A"
                        filled_columns_report.append(f"**{merge_filenames[i]}**: Wurde ergänzt um {missing_cols}")
                
                if filled_columns_report:
                    st.warning("⚠️ **Hinweis zur Struktur:** Nicht alle Dateien hatten dieselben Spalten. Fehlende Werte wurden automatisch mit 'N/A' aufgefüllt.")
                    with st.expander("Details zu ergänzten Spalten anzeigen"):
                        for report in filled_columns_report:
                            st.markdown(f"- {report}")

                # 2. Analyse auf Konflikte
                all_ids = set()
                all_meds = set()
                all_ops = set()
                all_music = set()
                
                for df in merge_dfs:
                    if 'participant_id' in df.columns:
                        all_ids.update(df['participant_id'].dropna().astype(str).unique())
                    if 'global_medikamente' in df.columns:
                        all_meds.update(df['global_medikamente'].dropna().astype(str).unique())
                    if 'global_operationen' in df.columns:
                        all_ops.update(df['global_operationen'].dropna().astype(str).unique())
                    if 'global_musikeffekt' in df.columns:
                        all_music.update(df['global_musikeffekt'].dropna().astype(str).unique())

                # 3. Validierung & UI für Konfliktlösung
                block_merge = False
                
                # A) Ranking Types Konflikt (BLOCKER)
                duplicate_conflicts = {rt: files for rt, files in ranking_type_sources.items() if len(files) > 1}
                
                if duplicate_conflicts:
                    st.error("⛔ **Konflikt: Doppelte Ranking-Typen erkannt!**")
                    st.markdown("Das Zusammenfügen ist nicht möglich, da folgende Typen in mehreren Dateien gleichzeitig vorkommen:")
                    for r_type, files in duplicate_conflicts.items():
                        st.markdown(f"- Ranking-Typ `{r_type}` gefunden in: **{', '.join(files)}**")
                    block_merge = True

                # B) Participant ID Konflikt
                final_id = list(all_ids)[0] if all_ids else "Unbekannt"
                if len(all_ids) > 1:
                    st.error(f"⚠️ **Konflikt: Unterschiedliche Teilnehmer-IDs gefunden:** {', '.join(all_ids)}")
                    final_id = st.selectbox("Welche ID soll für die gemergte Datei verwendet werden?", list(all_ids), key="sel_merge_id")
                
                # C) Meds & Ops Konflikt
                final_meds = list(all_meds)[0] if all_meds else "N/A"
                if len(all_meds) > 1:
                    st.error(f"⚠️ **Konflikt: Unterschiedliche Angaben zu Medikamenten:**")
                    final_meds = st.selectbox("Welche Angabe soll übernommen werden (Medikamente)?", list(all_meds), key="sel_merge_meds")

                final_ops = list(all_ops)[0] if all_ops else "N/A"
                if len(all_ops) > 1:
                    st.error(f"⚠️ **Konflikt: Unterschiedliche Angaben zu Operationen:**")
                    final_ops = st.selectbox("Welche Angabe soll übernommen werden (Operationen)?", list(all_ops), key="sel_merge_ops")

                # D) Music Effect Warnung
                final_music = list(all_music)[0] if all_music else "Nein"
                if len(all_music) > 1:
                    st.warning(f"⚠️ **Hinweis:** Unterschiedliche Angaben zum Musikeffekt gefunden ({', '.join(all_music)}).")
                    final_music = st.selectbox("Welche Angabe soll übernommen werden (Musikeffekt)?", list(all_music), key="sel_merge_music")

                # 4. Erstellung
                if not block_merge:
                    merged_df = pd.concat(merge_dfs, ignore_index=True)
                    
                    # Metadaten überschreiben
                    if 'participant_id' in merged_df.columns:
                        merged_df['participant_id'] = final_id
                    if 'global_medikamente' in merged_df.columns:
                        merged_df['global_medikamente'] = final_meds
                    if 'global_operationen' in merged_df.columns:
                        merged_df['global_operationen'] = final_ops
                    if 'global_musikeffekt' in merged_df.columns:
                        merged_df['global_musikeffekt'] = final_music
                    
                    # --- SORTIERLOGIK ---
                    sort_order = ['audio_1', 'audio_2', 'audio_3', 'audio_4', 'youtube']
                    order_map = {key: i for i, key in enumerate(sort_order)}
                    
                    # Temporäre Hilfsspalte für Sortierung
                    merged_df['__sort_helper'] = merged_df['ranking_type'].map(order_map).fillna(len(sort_order))
                    
                    # Sortieren und Hilfsspalte entfernen
                    merged_df = merged_df.sort_values(by=['__sort_helper', 'play_order']).drop(columns=['__sort_helper'])

                    # CSV Generierung
                    csv_data = merged_df.to_csv(index=False).encode('utf-8')
                    file_name_merged = f"export_{final_id}_merged.csv"
                    
                    st.markdown("###")
                    st.success(f"✅ Bereit zum Download! ({len(merged_df)} Einträge aus {len(merge_dfs)} Dateien, sortiert)")
                    
                    st.download_button(
                        label="📥 Zusammenfügen & Herunterladen",
                        data=csv_data,
                        file_name=file_name_merged,
                        mime='text/csv',
                        type="primary",
                        width='stretch'
                    )

    st.markdown("---"); st.subheader("2. EEG-Aufnahmen hochladen")
    
    # --- EEG UPLOADS ---
    
    # Baseline (Alleine)
    raw_bl = st.file_uploader("Grundaktivität (Baseline) EEG-Aufnahme", type="csv", key="up_bl")
    file_bl = persist_uploaded_file(raw_bl, 'stored_baseline')
    if file_bl and raw_bl is None: st.success(f"✅ Baseline geladen: {file_bl.name}")
    
    # YouTube (Alleine in einer "Zeile")
    raw_yt = st.file_uploader("YouTube Playlist EEG", type="csv", key="up_yt")
    file_yt = persist_uploaded_file(raw_yt, 'stored_youtube')
    if file_yt and raw_yt is None: st.success(f"✅ Datei geladen: {file_yt.name}")

    # Audio 1 & 2 (Nebeneinander)
    cols_audio_12 = st.columns(2)
    
    raw_a1 = cols_audio_12[0].file_uploader("Audio 1 Playlist EEG", type="csv", key="up_a1")
    file_a1 = persist_uploaded_file(raw_a1, 'stored_audio1')
    if file_a1 and raw_a1 is None: cols_audio_12[0].success(f"✅ Datei geladen: {file_a1.name}")

    raw_a2 = cols_audio_12[1].file_uploader("Audio 2 Playlist EEG", type="csv", key="up_a2")
    file_a2 = persist_uploaded_file(raw_a2, 'stored_audio2')
    if file_a2 and raw_a2 is None: cols_audio_12[1].success(f"✅ Datei geladen: {file_a2.name}")

    # Playlists 3 & 4 (Nebeneinander)
    cols_audio_34 = st.columns(2)
    
    raw_a3 = cols_audio_34[0].file_uploader("Audio 3 (Shuffled) Playlist EEG", type="csv", key="up_a3")
    file_a3 = persist_uploaded_file(raw_a3, 'stored_audio3')
    if file_a3 and raw_a3 is None: cols_audio_34[0].success(f"✅ Datei geladen: {file_a3.name}")

    raw_a4 = cols_audio_34[1].file_uploader("Audio 4 (Shuffled) Playlist EEG", type="csv", key="up_a4")
    file_a4 = persist_uploaded_file(raw_a4, 'stored_audio4')
    if file_a4 and raw_a4 is None: cols_audio_34[1].success(f"✅ Datei geladen: {file_a4.name}")

    # Mapping für Verarbeitung
    eeg_files_processing = {
        'youtube': file_yt, 
        'audio_1': file_a1, 
        'audio_2': file_a2,
        'audio_3': file_a3,
        'audio_4': file_a4
    }
    
    st.markdown("---"); st.subheader("3. Einstellungen")
    
    # Einstellungen nebeneinander (2 Spalten)
    col_set_1, col_set_2 = st.columns(2)
    
    with col_set_1:
        short_event_seconds = st.number_input(
            "Zeitfenster für 'kurze' Events (in Sekunden)", 0.5, 10.0, 2.5, 0.5,
            help="Wegschneidezeitraum in beide Richtungen für kurze Events (z.B. Husten)."
        )
        
    with col_set_2:
        reaction_buffer = st.number_input(
            "Reaktions-Puffer (in Sekunden)", 0.0, 5.0, 1.0, 0.5,
            help="Erweitert den Ausschnitt rückwirkend (in die Vergangenheit). Beispiel: Bei einem kurzen Event von 2,5s Radius und 1s Puffer wird 3,5s vor dem Trigger und 2,5s nach dem Trigger geschnitten. Bei langen Events wird der Startpunkt um diesen Wert vorverlegt."
        )

    st.markdown("---")
    
    eeg_dataframes = {}
    
    # --- DATEIEN LADEN UND VALIDIEREN ---
    
    # 1. Baseline
    if file_bl:
        try:
            file_bl.seek(0)
            df = pd.read_csv(file_bl)
            df.columns = df.columns.str.strip()
            err = validate_columns(df, REQUIRED_EEG_COLS, "Baseline")
            if err: validation_errors.append(err)
            else: eeg_dataframes['baseline'] = df
        except Exception as e: validation_errors.append(f"Fehler beim Lesen der Baseline-Datei: {e}")
    
    # 2. Playlists
    for r_type, file_obj in eeg_files_processing.items():
        if file_obj:
            try:
                file_obj.seek(0)
                df = pd.read_csv(file_obj)
                df.columns = df.columns.str.strip()
                err = validate_columns(df, REQUIRED_EEG_COLS, f"{r_type} Playlist")
                if err: validation_errors.append(err)
                else: eeg_dataframes[r_type] = df
            except Exception as e: validation_errors.append(f"Fehler beim Lesen der {r_type}-Datei: {e}")

    # --- ANZEIGELOGIK ---
    
    if eeg_dataframes:
        st.header("📋 Event-Zusammenfassung der Aufnahmen")
        valid_keys = list(eeg_dataframes.keys())
        num_cols = len(valid_keys)
        # Dynamisches Spalten-Layout
        if num_cols > 0:
            # Bei vielen Dateien brechen wir um
            cols_per_row = 3
            for i in range(0, num_cols, cols_per_row):
                row_keys = valid_keys[i:i+cols_per_row]
                event_cols = st.columns(len(row_keys))
                for j, name in enumerate(row_keys):
                    df = eeg_dataframes[name]
                    with event_cols[j]:
                        st.subheader(f"{name.replace('_', ' ').title()}")
                        status = None; is_baseline = (name == 'baseline')
                        if is_baseline:
                            trigger_col = 'TRIG'
                            if trigger_col in df.columns:
                                trig_vals = df[df[trigger_col] != 0][trigger_col].tolist()
                                start_count = trig_vals.count(1); end_count = trig_vals.count(2)
                                starts = df[df[trigger_col] == 1].index; ends = df[df[trigger_col] == 2].index
                                valid_order = False
                                if len(starts) == 1 and len(ends) == 1:
                                    if ends[0] > starts[0]: valid_order = True
                                if start_count == 1 and end_count == 1 and valid_order: status = "✅ Valide Baseline"
                                else: status = "❌ Invalide Baseline"
                        else: _, status = analyze_playlist_triggers(df)
                        display_event_counts(count_events(df), name, status, is_baseline=is_baseline)
        st.markdown("---")

    if validation_errors:
        st.error("⚠️ Die Analyse kann nicht gestartet werden, da folgende Fehler gefunden wurden:")
        for err in validation_errors: st.error(err)
    
    can_start = (not validation_errors) and ('baseline' in eeg_dataframes) and (len(eeg_dataframes) > 1) and (not user_ranking_df.empty)
    
    if can_start:
        if st.button("Analyse starten (Alle Reinigungs-Level)", type="primary", width='stretch'):
            results_collection = {}
            with st.spinner("Analysiere Daten für alle Reinigungs-Stufen..."):
                for cleaning_level in CLEANING_LEVELS:
                    all_playlist_results = {}
                    baseline_df = eeg_dataframes['baseline']
                    b_starts = baseline_df[baseline_df['TRIG'] == 1].index; b_ends = baseline_df[baseline_df['TRIG'] == 2].index
                    bl_long, bl_trans, bl_short = get_global_events(baseline_df)
                    raw_baseline_segment = baseline_df.iloc[b_starts[0]:b_ends[0]]
                    
                    cleaned_baseline_segment = clean_segment_with_context(raw_baseline_segment, cleaning_level, short_event_seconds, reaction_buffer, bl_long, bl_trans, bl_short)

                    if cleaned_baseline_segment.empty:
                        baseline_result = {'Phase': 'Grundaktivität', 'Alpha/Beta Verhältnis': 0.0, 'Theta/Beta Verhältnis': 0.0, '(Alpha+Theta)/Beta Verhältnis': 0.0}
                    else:
                        baseline_result = {'Phase': 'Grundaktivität', **calculate_relaxation_ratios(compute_band_power(cleaned_baseline_segment, SAMPLING_FREQUENCY, EEG_CHANNELS))}
                    
                    for r_type, df_playlist in eeg_dataframes.items():
                        if r_type == 'baseline': continue
                        global_long, global_trans, global_short = get_global_events(df_playlist)
                        song_segments, _ = analyze_playlist_triggers(df_playlist)
                        if not song_segments: continue
                        chrono_names, pref_names = get_rankings_from_df(user_ranking_df, r_type)
                        if not pref_names: continue
                        
                        eeg_results = []
                        eeg_results.append(baseline_result)
                        for i, segment in enumerate(song_segments):
                            phase_name = chrono_names[i] if i < len(chrono_names) else f'Item {i+1} (Unbekannt)'
                            
                            cleaned_segment = clean_segment_with_context(segment, cleaning_level, short_event_seconds, reaction_buffer, global_long, global_trans, global_short)
                            
                            if cleaned_segment.empty: continue
                            eeg_results.append({'Phase': phase_name, **calculate_relaxation_ratios(compute_band_power(cleaned_segment, SAMPLING_FREQUENCY, EEG_CHANNELS))})
                        all_playlist_results[r_type] = {'preference_ranking': pref_names, 'eeg_results_df': pd.DataFrame(eeg_results)}
                    results_collection[cleaning_level] = all_playlist_results
                st.session_state.analysis_results = results_collection
    
    elif not validation_errors and (not 'baseline' in eeg_dataframes or len(eeg_dataframes) <= 1):
        st.info("Lade mindestens die Baseline und eine Playlist-Datei hoch, um die Analyse, mithilfe des User-Rankings, starten zu können.")

    if st.session_state.analysis_results:
        results_collection = st.session_state.analysis_results
        st.header("Gesamtanalyse aller Playlists")
        
        # Tabs für Reinigungs-Level
        level_tabs = st.tabs(CLEANING_LEVELS)
        
        for i, cleaning_level in enumerate(CLEANING_LEVELS):
            with level_tabs[i]:
                all_playlist_results = results_collection.get(cleaning_level, {})
                if not all_playlist_results: 
                    st.warning(f"Keine Ergebnisse für Level '{cleaning_level}'.")
                    continue

                # --- Playlist Tabs ---
                playlist_keys = list(all_playlist_results.keys())
                playlist_labels = [k.replace('_', ' ').title() for k in playlist_keys]
                
                if not playlist_labels: continue
                
                pl_tabs = st.tabs(playlist_labels)
                
                for idx, pl_key in enumerate(playlist_keys):
                    data = all_playlist_results[pl_key]
                    with pl_tabs[idx]:
                        results_df = data['eeg_results_df']
                        preference_item_names = data['preference_ranking'] 
                        
                        if results_df.empty: 
                            st.warning("Keine EEG-Daten.")
                            continue
                        
                        eeg_items_set = set(results_df['Phase'])
                        common_items_user_order = [item for item in preference_item_names if item in eeg_items_set]
                        
                        if not common_items_user_order:
                            st.warning("Keine Übereinstimmung zwischen User-Ranking und EEG-Daten.")
                            continue

                        metrics = ['Alpha/Beta Verhältnis', 'Theta/Beta Verhältnis', '(Alpha+Theta)/Beta Verhältnis']
                        num_items = len(common_items_user_order)
                        
                        # --- ERWEITERTE LOGIK-FUNKTION ---
                        def calculate_tab_content(use_optimization):
                            """
                            Berechnet Scores und Tabelle.
                            use_optimization = False -> Striktes Sortieren nach Messwert.
                            use_optimization = True  -> Sortieren mit Toleranz (Threshold) zugunsten des Users.
                            """
                            
                            total_match_ranks = 0 
                            individual_matches = {m: 0 for m in metrics}
                            theoretical_max = num_items
                            
                            comp_data = {
                                'Rang': range(1, num_items + 1), 
                                'Dein Ranking': common_items_user_order
                            }
                            
                            row_eeg_items = {m: [] for m in metrics}
                            
                            # 1. Listen vorbereiten
                            eeg_sorted_lists = {}
                            for m in metrics:
                                vals = results_df[m].values
                                thresh = calculate_dynamic_threshold(vals, factor=0.5)
                                
                                if use_optimization:
                                    # Optimiert: Nutzt Threshold zum "Cluster-Tausch"
                                    eeg_sorted_lists[m] = get_sehr_fair_ranking_order(results_df, m, common_items_user_order, thresh)
                                else:
                                    # Strikt: Ignoriert Threshold, sortiert rein numerisch
                                    eeg_sorted_lists[m] = get_strict_ranking_order(results_df, m)
                                
                                eeg_sorted_lists[m] = [x for x in eeg_sorted_lists[m] if x in common_items_user_order]

                            # 2. Iteration über die Ränge
                            for i in range(num_items):
                                user_item = common_items_user_order[i]
                                row_has_any_match = False
                                suggestions_at_rank = []
                                
                                for m in metrics:
                                    eeg_list = eeg_sorted_lists[m]
                                    if i < len(eeg_list):
                                        eeg_item = eeg_list[i]
                                        suggestions_at_rank.append(eeg_item)
                                        row_eeg_items[m].append(eeg_item)
                                        
                                        if eeg_item == user_item:
                                            individual_matches[m] += 1
                                            row_has_any_match = True
                                    else:
                                        row_eeg_items[m].append("-")
                                
                                if row_has_any_match:
                                    total_match_ranks += 1
                                
                                unique_suggestions = set(suggestions_at_rank)
                                bonus = len(suggestions_at_rank) - len(unique_suggestions)
                                theoretical_max += bonus

                            dominant = determine_dominant_metric(individual_matches)
                            interpretations = {
                                'Alpha/Beta Verhältnis': "Du assoziierst Entspannung am ehesten mit einem Zustand **wacher Aufmerksamkeit**.", 
                                'Theta/Beta Verhältnis': "Du assoziierst Entspannung am ehesten mit einem **tiefen, meditativen Zustand**, der an Schläfrigkeit grenzt.", 
                                '(Alpha+Theta)/Beta Verhältnis': "Du assoziierst Entspannung am ehesten mit einer **ausgeglichenen Mischung** aus wacher und tiefer Ruhe."
                            }
                            interp_text = interpretations.get(dominant, "Keine klare Tendenz.") if dominant else "Zu wenige Treffer."

                            for m in metrics:
                                short_m = m.replace(' Verhältnis', '').replace('Alpha', 'A').replace('Theta', 'T').replace('Beta', 'B')
                                comp_data[f"EEG {short_m}"] = row_eeg_items[m]
                                
                            df_res = pd.DataFrame(comp_data).set_index('Rang')
                            
                            stats = {
                                'total_ranks': total_match_ranks,
                                'total_max': theoretical_max,
                                'green_score': sum(individual_matches.values()),
                                'individual': individual_matches,
                                'interpretation': interp_text
                            }
                            return df_res, stats

                        # --- UI START ---
                        
                        # HIER WURDEN DIE NAMEN ANGEPASST:
                        ana_tab1, ana_tab2 = st.tabs(["Strikt (Numerisch)", "Optimiert (Mit Toleranz)"])
                        
                        def display_score_block(stats_dict):
                            col1, col2 = st.columns(2)
                            col1.metric("Gesamt-Trefferquote (Zeilen)", f"{stats_dict['total_ranks']} / {num_items}")
                            col2.metric("Punktzahl vs. theoretische Max.", f"{stats_dict['green_score']} / {stats_dict['total_max']}")
                            st.markdown(f"**Interpretation:** {stats_dict['interpretation']}")
                            
                            st.caption("Detaillierte Treffer pro Metrik:")
                            c_s1, c_s2, c_s3 = st.columns(3)
                            indiv = stats_dict['individual']
                            c_s1.metric("Alpha/Beta", f"{indiv['Alpha/Beta Verhältnis']}/{num_items}")
                            c_s2.metric("Theta/Beta", f"{indiv['Theta/Beta Verhältnis']}/{num_items}")
                            c_s3.metric("(A+T)/Beta", f"{indiv['(Alpha+Theta)/Beta Verhältnis']}/{num_items}")
                            st.divider()

                        def highlight_fair(x):
                            df = pd.DataFrame('', index=x.index, columns=x.columns)
                            user_col = x['Dein Ranking']
                            for col in x.columns:
                                if col.startswith('EEG'):
                                    mask = x[col] == user_col
                                    df[col] = np.where(mask, 'background-color: #28a745; color: white', 'background-color: #dc3545; color: white')
                            return df

                        # --- TAB 1: STRIKT ---
                        with ana_tab1:
                            df_strict, stats_strict = calculate_tab_content(use_optimization=False)
                            display_score_block(stats_strict)
                            st.caption("Vergleich: **Strikte Logik**. Sortierung rein nach Messwert (höchster gewinnt). Keine Toleranz für knappe Unterschiede.")
                            st.dataframe(df_strict.style.apply(highlight_fair, axis=None), width='stretch')

                        # --- TAB 2: OPTIMIERT ---
                        with ana_tab2:
                            df_opt, stats_opt = calculate_tab_content(use_optimization=True)
                            
                            delta_ranks = stats_opt['total_ranks'] - stats_strict['total_ranks']
                            display_score_block(stats_opt)
                            if delta_ranks > 0:
                                st.success(f"📈 Durch Berücksichtigung des Thresholds (Toleranz) wurden **{delta_ranks}** zusätzliche Übereinstimmungen gefunden (physiologisch gleichwertige Items getauscht).")
                            
                            st.caption("Vergleich: **Optimierte Logik**. Werte innerhalb des Thresholds gelten als gleichwertig und werden so sortiert, dass sie deinem Ranking entsprechen.")
                            st.dataframe(df_opt.style.apply(highlight_fair, axis=None), width='stretch')

                        # --- EXPANDER: ROHDATEN ---
                        st.markdown("###")
                        with st.expander("📄 Konsolidiertes EEG Ranking (Rohwerte anzeigen)", expanded=False):
                            st.info("ℹ️ **Berechnung des Thresholds (Toleranzwert):**\n"
                                    "Der Wert berechnet sich dynamisch pro Verhältnis aus: "
                                    "**Standardabweichung (aller Werte inkl. Baseline) × 0,5**.")
                            
                            ranking_data = {}
                            for m in metrics:
                                vals = results_df[m].values
                                thresh = calculate_dynamic_threshold(vals, factor=0.5)
                                sorted_df = results_df.sort_values(by=m, ascending=False)
                                
                                short_m_name = m.replace(' Verhältnis', '')
                                col_name = f"{short_m_name}\n(Threshold: ±{thresh:.4f})"
                                
                                ranking_data[col_name] = [
                                    f"**{row['Phase']}** ({row[m]:.4f})" 
                                    for _, row in sorted_df.iterrows()
                                ]
                            
                            max_len = max(len(v) for v in ranking_data.values())
                            for k in ranking_data:
                                ranking_data[k] += [""] * (max_len - len(ranking_data[k]))
                                
                            df_raw = pd.DataFrame(ranking_data)
                            df_raw.index = range(1, len(df_raw) + 1)
                            st.table(df_raw)

        st.markdown("---")
        st.subheader("💾 Gesamt-Report exportieren")
        
        date_str = datetime.now().strftime('%d-%m-%y')
        safe_id = participant_id if participant_id != "Unbekannt" else "EEG_Analyse"
        export_filename = f"{safe_id}_{date_str}.html"
        
        html_report = generate_html_report(results_collection, short_event_seconds, reaction_buffer, participant_id)
        
        st.download_button(
            label="📄 Ergebnisse als HTML-Report herunterladen", 
            data=html_report, 
            file_name=export_filename, 
            mime="text/html", 
            type="primary",
            width='stretch'
        )