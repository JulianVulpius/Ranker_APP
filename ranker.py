import streamlit as st
import os
from googleapiclient.discovery import build
import re
import isodate
import streamlit_antd_components as sac
from ranker_component import youtube_ranker, audio_ranker
from datetime import datetime
import base64
import pandas as pd
import io

from auswertung_seite import display_auswertung_page

st.set_page_config(layout="wide")

# --- KONFIGURATION & KONSTANTEN ---

BASELINE_ITEM_AUDIO = {"id": "baseline", "name": "Grundaktivität", "image_path": "assets/images/baseline.jpg", "audio_path": "", "bekannt": False, "gefallen": "Neutral", "entspannt": False}
BASELINE_ITEM_YOUTUBE = {'id': 'baseline', 'title': 'Grundaktivität', 'thumbnail': 'static/assets/images/baseline.jpg', 'length': None, 'entspannt': False}

# Erweiterte Ranking Typen
RANKING_TYPES = ['youtube', 'audio_1', 'audio_2', 'audio_3', 'audio_4']
QUESTIONS = ['entspannung', 'wohlbefinden', 'muedigkeit']

# State Initialisierung für Fragen
for r_type in RANKING_TYPES:
    for question in QUESTIONS:
        key = f"{r_type}_{question}"
        if key not in st.session_state:
            st.session_state[key] = 5

# --- HELPER FUNKTION FÜR ITEM ERSTELLUNG ---
def create_song_item(i, play_order, is_shuffled=False, display_name=None):
    """Erstellt ein Song-Item Dictionary."""
    item = {
        "id": f"song_{i}", 
        "name": f"Song {i}", 
        "image_path": f"assets/images/song{i}.jpg", 
        "audio_path": f"assets/audio/song{i}.mp3", 
        "bekannt": True if is_shuffled else False,
        "gefallen": "Neutral", 
        "entspannt": False, 
        "play_order": play_order
    }
    if display_name:
        item["display_name"] = display_name
    return item

# --- INITIALISIERUNG DER PLAYLISTS IM SESSION STATE ---

if 'ranked_audio_1' not in st.session_state:
    AUDIO_ITEMS_DATA_1 = [create_song_item(i, play_order=i) for i in range(1, 6)]
    baseline_item_1 = BASELINE_ITEM_AUDIO.copy()
    baseline_item_1['play_order'] = 0
    AUDIO_ITEMS_DATA_1.insert(0, baseline_item_1)
    st.session_state['ranked_audio_1'] = AUDIO_ITEMS_DATA_1

if 'ranked_audio_2' not in st.session_state:
    AUDIO_ITEMS_DATA_2 = [create_song_item(i, play_order=i-5) for i in range(6, 11)]
    baseline_item_2 = BASELINE_ITEM_AUDIO.copy()
    baseline_item_2['play_order'] = 0
    AUDIO_ITEMS_DATA_2.insert(0, baseline_item_2)
    st.session_state['ranked_audio_2'] = AUDIO_ITEMS_DATA_2

# Mapping: song5, song3, song8, song2, song6
# Display Namen: Song 1 bis Song 5
if 'ranked_audio_3' not in st.session_state:
    # Tuple Format: (Original ID Number, Play Order/Display Number)
    mapping_3 = [(5, 1), (3, 2), (8, 3), (2, 4), (6, 5)]
    AUDIO_ITEMS_DATA_3 = []
    
    for orig_id, display_num in mapping_3:
        item = create_song_item(orig_id, play_order=display_num, is_shuffled=True, display_name=f"Song {display_num}")
        AUDIO_ITEMS_DATA_3.append(item)
        
    baseline_item_3 = BASELINE_ITEM_AUDIO.copy()
    baseline_item_3['play_order'] = 0
    baseline_item_3['bekannt'] = True # Auch Baseline auf bekannt setzen bei Shuffled
    AUDIO_ITEMS_DATA_3.insert(0, baseline_item_3)
    st.session_state['ranked_audio_3'] = AUDIO_ITEMS_DATA_3

# Mapping: song10, song4, song9, song1, song7
# Display Namen: Song 6 bis Song 10
if 'ranked_audio_4' not in st.session_state:
    mapping_4 = [(10, 1, 6), (4, 2, 7), (9, 3, 8), (1, 4, 9), (7, 5, 10)] # (OrigID, PlayOrder, DisplayNum)
    AUDIO_ITEMS_DATA_4 = []
    
    for orig_id, p_order, d_num in mapping_4:
        item = create_song_item(orig_id, play_order=p_order, is_shuffled=True, display_name=f"Song {d_num}")
        AUDIO_ITEMS_DATA_4.append(item)
        
    baseline_item_4 = BASELINE_ITEM_AUDIO.copy()
    baseline_item_4['play_order'] = 0
    baseline_item_4['bekannt'] = True
    AUDIO_ITEMS_DATA_4.insert(0, baseline_item_4)
    st.session_state['ranked_audio_4'] = AUDIO_ITEMS_DATA_4


# --- UTILS ---

def get_base64_data_url(file_path):
    try:
        with open(file_path, "rb") as f: file_bytes = f.read()
        base64_encoded_data = base64.b64encode(file_bytes).decode('utf-8')
        return f"data:image/jpeg;base64,{base64_encoded_data}"
    except FileNotFoundError: return None

def get_playlist_videos(api_key: str, playlist_id: str) -> list:
    youtube = build('youtube', 'v3', developerKey=api_key)
    playlist_request = youtube.playlistItems().list(part='contentDetails', playlistId=playlist_id, maxResults=5)
    playlist_response = playlist_request.execute()
    video_ids = [item['contentDetails']['videoId'] for item in playlist_response.get('items', [])]
    if not video_ids: return []
    video_request = youtube.videos().list(part='snippet,contentDetails', id=','.join(video_ids))
    video_response = video_request.execute()
    videos_data = []
    video_details_map = {item['id']: item for item in video_response.get('items', [])}
    for i, video_id in enumerate(video_ids):
        item = video_details_map.get(video_id)
        if item:
            duration_iso = item['contentDetails']['duration']
            duration_seconds = isodate.parse_duration(duration_iso).total_seconds()
            minutes, seconds = divmod(duration_seconds, 60)
            duration_formatted = f"{int(minutes):02}:{int(seconds):02}"
            videos_data.append({
                'id': item['id'], 'title': item['snippet']['title'],
                'thumbnail': item['snippet']['thumbnails']['high']['url'],
                'duration_seconds': duration_seconds,
                'length': duration_formatted, 'play_order': i + 1,
                'entspannt': False 
            })
    baseline_item_yt = BASELINE_ITEM_YOUTUBE.copy()
    baseline_item_yt['play_order'] = 0
    videos_data.insert(0, baseline_item_yt)
    return videos_data

def get_playlist_id(url: str) -> str | None:
    pattern = r'(?:https?:\/\/)?(?:www.)?youtube\.com\/playlist\?list=([a-zA-Z0-9_-]+)'
    match = re.search(pattern, url)
    return match.group(1) if match else None

# --- EXPORT FUNKTIONEN ---

def generate_txt_export():
    export_data = []
    participant_id = st.session_state.get('participant_id', '').strip()
    export_data.append(f"Teilnehmer ID: {participant_id if participant_id else 'N/A'}")
    export_data.append(f"Export Datum: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    export_data.append("\n--- Allgemeine Antworten ---")
    medikamente = st.session_state.get('global_medikamente', '').strip()
    export_data.append(f"Heute irgendwelche Medikamente eingenommen? -> {medikamente if medikamente else 'N/A'}")
    operationen = st.session_state.get('global_operationen', '').strip()
    export_data.append(f"Wurden Operationen am Schädel gemacht? (Wenn ja, welche?) -> {operationen if operationen else 'N/A'}")
    export_data.append(f"Ich kann mich besser Entspannen ohne Musik -> {'Ja' if st.session_state.get('global_musikeffekt', False) else 'Nein'}")

    # Helper function for exporting sections
    def append_ranking_section(title, key_prefix, data_key):
        export_data.append("\n" + "="*20 + f"\n{title}\n" + "="*20)
        export_data.append(f"\n--- Antworten ({title}) ---\nWie entspannt bist du von 1-10? -> {st.session_state.get(f'{key_prefix}_entspannung', 'N/A')}")
        export_data.append(f"Wie gut geht es dir von 1-10? -> {st.session_state.get(f'{key_prefix}_wohlbefinden', 'N/A')}")
        export_data.append(f"Wie müde bist du von 1-10? -> {st.session_state.get(f'{key_prefix}_muedigkeit', 'N/A')}")
        export_data.append(f"\n--- Finales Ranking ({title}) ---")
        ranked_items = st.session_state.get(data_key, [])
        if ranked_items:
            for i, item in enumerate(ranked_items):
                entspannt_str = "Ja" if item.get('entspannt', False) else "Nein"
                # Use 'title' for YT, 'name' for Audio. Export always uses original name!
                name = item.get('title') or item.get('name', 'N/A')
                duration_str = f" (Dauer: {item.get('length', 'N/A')})" if item.get('length') else ""
                
                if item['id'] == 'baseline': 
                    export_data.append(f"{i+1}. {name}{duration_str} | Entspannt: {entspannt_str}")
                else: 
                    bekannt_str = 'Ja' if item.get('bekannt', False) else 'Nein'
                    gefallen_str = item.get('gefallen', 'N/A')
                    export_data.append(f"{i+1}. {name}{duration_str} | Bekannt: {bekannt_str} | Gefallen: {gefallen_str} | Entspannt: {entspannt_str}")
        else: export_data.append("Ranking nicht durchgeführt: N/A")

    if st.session_state.get('export_youtube', False):
        append_ranking_section("YOUTUBE RANKING", "youtube", "ranked_videos")

    if st.session_state.get('export_audio_1', False):
        append_ranking_section("AUDIO RANKING 1", "audio_1", "ranked_audio_1")

    if st.session_state.get('export_audio_2', False):
        append_ranking_section("AUDIO RANKING 2", "audio_2", "ranked_audio_2")
        
    if st.session_state.get('export_audio_3', False):
        append_ranking_section("AUDIO RANKING 3 (SHUFFLED)", "audio_3", "ranked_audio_3")

    if st.session_state.get('export_audio_4', False):
        append_ranking_section("AUDIO RANKING 4 (SHUFFLED)", "audio_4", "ranked_audio_4")

    return "\n".join(export_data)

def generate_csv_export():
    all_rows = []
    participant_id = st.session_state.get('participant_id', 'N/A').strip()
    export_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    general_data = {
        'participant_id': participant_id, 'export_timestamp': export_timestamp,
        'global_medikamente': st.session_state.get('global_medikamente', 'N/A').strip(),
        'global_operationen': st.session_state.get('global_operationen', 'N/A').strip(),
        'global_musikeffekt': 'Ja' if st.session_state.get('global_musikeffekt', False) else 'Nein'
    }
    ranking_map = {
        'youtube': ('export_youtube', 'ranked_videos'),
        'audio_1': ('export_audio_1', 'ranked_audio_1'),
        'audio_2': ('export_audio_2', 'ranked_audio_2'),
        'audio_3': ('export_audio_3', 'ranked_audio_3'),
        'audio_4': ('export_audio_4', 'ranked_audio_4'),
    }
    for r_type, (export_key, state_key) in ranking_map.items():
        if st.session_state.get(export_key, False):
            ranking_data = st.session_state.get(state_key, [])
            ranking_questions = {
                f'ranking_entspannung': st.session_state.get(f'{r_type}_entspannung', 'N/A'),
                f'ranking_wohlbefinden': st.session_state.get(f'{r_type}_wohlbefinden', 'N/A'),
                f'ranking_muedigkeit': st.session_state.get(f'{r_type}_muedigkeit', 'N/A'),
            }
            for i, item in enumerate(ranking_data):
                row = general_data.copy()
                row.update(ranking_questions)
                row.update({
                    'ranking_type': r_type, 'final_rank': i + 1,
                    'play_order': item.get('play_order', 'N/A'),
                    'item_id': item.get('id', 'N/A'),
                    # EXPORTIERT DEN ORIGINAL NAMEN, AUCH WENN DISPLAY NAME EXISTIERT
                    'item_name': item.get('title') or item.get('name', 'N/A'),
                    'item_duration': item.get('length', 'N/A'),
                    'item_bekannt': 'Ja' if item.get('bekannt', False) else 'Nein' if 'bekannt' in item else 'N/A',
                    'item_gefallen': item.get('gefallen', 'N/A'),
                    'item_entspannt': 'Ja' if item.get('entspannt', False) else 'Nein'
                })
                all_rows.append(row)
    if not all_rows: return ""
    df = pd.DataFrame(all_rows)
    cols_order = ['participant_id', 'export_timestamp', 'ranking_type', 'play_order', 'final_rank', 'item_name', 'item_id', 'item_duration', 'item_entspannt', 'item_bekannt', 'item_gefallen', 'ranking_entspannung', 'ranking_wohlbefinden', 'ranking_muedigkeit', 'global_medikamente', 'global_operationen', 'global_musikeffekt']
    df = df[[col for col in cols_order if col in df.columns]]
    output = io.StringIO()
    df.to_csv(output, index=False, encoding='utf-8')
    return output.getvalue()

def generate_automation_csv():
    videos_from_state = st.session_state.get('ranked_videos', [])
    if not videos_from_state:
        return ""

    sorted_videos = sorted(videos_from_state, key=lambda v: v.get('play_order', 999))
    
    export_data = []
    for video in sorted_videos:
        if video.get('id') == 'baseline':
            continue
        
        song_name = video.get('title', 'N/A')
        order = video.get('play_order', 'N/A') 
        duration_seconds = int(video.get('duration_seconds', 0))
        
        export_data.append([song_name, order, duration_seconds])
        
    df = pd.DataFrame(export_data, columns=['SongName', 'Reihenfolge', 'DauerInSekunden'])
    output = io.StringIO()
    df.to_csv(output, index=False, header=True, sep=',', quotechar='"')
    return output.getvalue()

# GLOBAL-UI
st.title("Entspannungs Ranker")
st.subheader("Allgemeine Informationen")
st.text_input(label="Teilnehmer ID", label_visibility="collapsed", value=st.session_state.get('participant_id', ''), key='participant_id', placeholder="Teilnehmer ID")
st.text_input(label="Heute irgendwelche Medikamente eingenommen?", label_visibility="collapsed", value=st.session_state.get('global_medikamente', ''), key='global_medikamente', placeholder="Heute irgendwelche Medikamente eingenommen? (Wenn ja, welche?)")
st.text_input(label="Wurden Operationen am Schädel gemacht?", label_visibility="collapsed", value=st.session_state.get('global_operationen', ''), key='global_operationen', placeholder="Wurden Operationen am Schädel gemacht? (Wenn ja, welche?)")
st.checkbox(label='Ich kann mich besser Entspannen ohne Musik', value=st.session_state.get('global_musikeffekt', False), key='global_musikeffekt')

selected_tab = sac.tabs([
    sac.TabsItem(label='Dateneingabe & Ranking', icon='pencil-square'),
    sac.TabsItem(label='Auswertung', icon='graph-up'),
], align='center', format_func='title')

# RANKING UI
if selected_tab == 'Dateneingabe & Ranking':
    st.title("🏆 Ranking")
    st.markdown("---")
    st.subheader("Export-Optionen")
    col1, col2 = st.columns(2)
    with col1:
        st.checkbox("YouTube-Daten", key='export_youtube', value=True)
        st.checkbox("Audio 1-Daten", key='export_audio_1', value=True)
        st.checkbox("Audio 2-Daten", key='export_audio_2', value=True)
        st.checkbox("Audio 3-Daten (Shuffled)", key='export_audio_3', value=True)
        st.checkbox("Audio 4-Daten (Shuffled)", key='export_audio_4', value=True)
    with col2:
        export_format = st.selectbox("Exportformat wählen", ("CSV (für Datenanalyse)", "TXT (für schnelle Lesbarkeit)"))

    if export_format == "TXT (für schnelle Lesbarkeit)":
        file_extension, mime_type, export_content = "txt", "text/plain", generate_txt_export()
    else:
        file_extension, mime_type, export_content = "csv", "text/csv", generate_csv_export()
    st.download_button(
        label="Daten exportieren", data=export_content,
        file_name=f"export_{st.session_state.get('participant_id', 'NO_ID')}_{datetime.now().strftime('%Y%m%d')}.{file_extension}",
        mime=mime_type, help="Klicken, um die ausgewählten Daten im gewählten Format herunterzuladen.", width='stretch'
    )
    st.markdown("---")
    ranking_tab = sac.tabs([
        sac.TabsItem(label='YouTube Ranker', icon='youtube'),
        sac.TabsItem(label='Audio Ranker 1', icon='music-note-beamed'),
        sac.TabsItem(label='Audio Ranker 2', icon='music-note-beamed'),
        sac.TabsItem(label='Audio Ranker 3', icon='shuffle'),
        sac.TabsItem(label='Audio Ranker 4', icon='shuffle'),
    ], align='center', format_func='title')

    def page_specific_questions(ranking_type: str):
        st.markdown("---"); st.subheader("Fragen vor dem Start der Playlist/Rankings")
        for question in QUESTIONS:
            key = f"{ranking_type}_{question}"
            question_text = {"entspannung": "Wie entspannt bist du von 1-10?", "wohlbefinden": "Wie gut geht es dir von 1-10?", "muedigkeit": "Wie müde bist du von 1-10?"}.get(question)
            st.markdown(f"<p style='font-size: 20px; font-weight: normal;'>{question_text}</p>", unsafe_allow_html=True)
            st.session_state[key] = st.number_input(
                label=question_text, label_visibility="collapsed", min_value=1, max_value=10,
                value=st.session_state[key], step=1, key=f"input_{key}"
            )

    # HELPER FÜR AUDIO UPDATE
    def update_audio_state(state_key, component_return):
        if component_return:
            current_sig = [(item['id'], item.get('bekannt', False), item.get('gefallen', 'Neutral'), item.get('entspannt', False)) for item in st.session_state[state_key]]
            returned_sig = [(item['id'], item.get('bekannt', False), item.get('gefallen', 'Neutral'), item.get('entspannt', False)) for item in component_return]
            if current_sig != returned_sig:
                orig_map = {item['id']: item for item in st.session_state[state_key]}
                reconstructed = []
                for ret_item in component_return:
                    item_id = ret_item.get('id')
                    if item_id in orig_map:
                        new_item = orig_map[item_id].copy()
                        new_item['bekannt'] = ret_item.get('bekannt', False)
                        new_item['gefallen'] = ret_item.get('gefallen', 'Neutral')
                        new_item['entspannt'] = ret_item.get('entspannt', False)
                        reconstructed.append(new_item)
                if reconstructed and len(reconstructed) == len(st.session_state[state_key]):
                    st.session_state[state_key] = reconstructed
                    st.rerun()

    if ranking_tab == 'YouTube Ranker':
        col_title, col_button = st.columns([13, 1])
        with col_title:
            st.header("YouTube Video Ranker")
        playlist_url = st.text_input("Geben Sie die URL einer öffentlichen YouTube-Playlist ein", value=st.session_state.get('playlist_url', ''))

        if playlist_url != st.session_state.get('playlist_url', ''):
            st.session_state.playlist_url = playlist_url
            playlist_id = get_playlist_id(playlist_url)
            if playlist_id:
                try:
                    API_KEY = st.secrets["API_KEY"]
                    with st.spinner("Videos werden abgerufen..."): st.session_state['ranked_videos'] = get_playlist_videos(API_KEY, playlist_id)
                except Exception as e: st.error(f"Ein Fehler ist aufgetreten: {e}"); st.session_state['ranked_videos'] = []
            else:
                if playlist_url: st.warning("Bitte geben Sie eine gültige YouTube-Playlist-URL ein.")
                st.session_state['ranked_videos'] = []
            st.rerun()

        if 'ranked_videos' in st.session_state and st.session_state['ranked_videos']:
            with col_button:
                st.write("") 
                st.write("")
                automation_csv_data = generate_automation_csv()
                st.download_button(
                    label="📥 Trigger CSV",
                    data=automation_csv_data,
                    file_name=f"automation_export_{st.session_state.get('participant_id', 'NO_ID')}_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    help="Exportiert die aktuelle Rangliste der YouTube-Videos als CSV für die Automatisierung."
                )

        if 'ranked_videos' in st.session_state and st.session_state['ranked_videos']:
            page_specific_questions('youtube')
            st.subheader("Songs per Drag-and-Drop in die gewünschte Reihenfolge bringen. [1 Platz = entspannenste Zustand]")
            processed_videos = []
            for video in st.session_state['ranked_videos']:
                if video['id'] == 'baseline' and os.path.exists(video['thumbnail']):
                    processed_videos.append({**video, 'thumbnail': get_base64_data_url(video['thumbnail'])})
                else: processed_videos.append(video)
            
            reordered_items = youtube_ranker(processed_videos, key='video_ranker')
            
            if reordered_items:
                current_state = st.session_state['ranked_videos']
                has_changes = False
                if len(reordered_items) != len(current_state): has_changes = True
                else:
                    for i, item in enumerate(reordered_items):
                        old_item = current_state[i]
                        if item['id'] != old_item['id'] or item.get('entspannt') != old_item.get('entspannt'):
                            has_changes = True; break
                if has_changes:
                    updated_list = []
                    orig_map = {item['id']: item for item in current_state}
                    for ret_item in reordered_items:
                        orig = orig_map.get(ret_item['id'])
                        if orig:
                            new_obj = orig.copy()
                            new_obj['entspannt'] = ret_item.get('entspannt', False)
                            updated_list.append(new_obj)
                    st.session_state['ranked_videos'] = updated_list; st.rerun()

    elif ranking_tab == 'Audio Ranker 1':
        st.header("Audio Ranker 1")
        page_specific_questions('audio_1')
        st.subheader("Songs per Drag-and-Drop in die gewünschte Reihenfolge bringen. [1 Platz = entspannenste Zustand]")
        ret = audio_ranker(st.session_state['ranked_audio_1'], key='audio_ranker_1')
        update_audio_state('ranked_audio_1', ret)

    elif ranking_tab == 'Audio Ranker 2':
        st.header("Audio Ranker 2")
        page_specific_questions('audio_2')
        st.subheader("Songs per Drag-and-Drop in die gewünschte Reihenfolge bringen. [1 Platz = entspannenste Zustand]")
        ret = audio_ranker(st.session_state['ranked_audio_2'], key='audio_ranker_2')
        update_audio_state('ranked_audio_2', ret)
        
    elif ranking_tab == 'Audio Ranker 3':
        st.header("Audio Ranker 3")
        page_specific_questions('audio_3')
        st.subheader("Songs per Drag-and-Drop in die gewünschte Reihenfolge bringen. [1 Platz = entspannenste Zustand]")
        ret = audio_ranker(st.session_state['ranked_audio_3'], key='audio_ranker_3')
        update_audio_state('ranked_audio_3', ret)

    elif ranking_tab == 'Audio Ranker 4':
        st.header("Audio Ranker 4")
        page_specific_questions('audio_4')
        st.subheader("Songs per Drag-and-Drop in die gewünschte Reihenfolge bringen. [1 Platz = entspannenste Zustand]")
        ret = audio_ranker(st.session_state['ranked_audio_4'], key='audio_ranker_4')
        update_audio_state('ranked_audio_4', ret)

elif selected_tab == 'Auswertung':
    display_auswertung_page()