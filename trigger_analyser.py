import streamlit as st
import pandas as pd

# --------------------------
# KONFIGURATION & STYLES
# --------------------------
st.set_page_config(page_title="Trigger Master Analyse", layout="wide")

TRIGGER_DEFS = {
    1: {"grp": "Song", "icon": "🎵", "name": "Start", "style": "background-color: #d1e7dd; color: #0f5132; border: 1px solid #badbcc;"}, 
    2: {"grp": "Song", "icon": "🎵", "name": "Ende",  "style": "background-color: #198754; color: white; border: 1px solid #198754;"},
    
    3: {"grp": "Visuell", "icon": "👁️", "name": "Visuell Start", "style": "background-color: #e2d9f3; color: #4a148c; border: 1px solid #d1c4e9;"},
    4: {"grp": "Visuell", "icon": "👁️", "name": "Visuell Ende",  "style": "background-color: #7b1fa2; color: white; border: 1px solid #7b1fa2;"},

    5: {"grp": "Auditiv", "icon": "👂", "name": "Start", "style": "background-color: #ffccbc; color: #bf360c; border: 1px solid #ffab91;"},
    6: {"grp": "Auditiv", "icon": "👂", "name": "Ende",  "style": "background-color: #d84315; color: white; border: 1px solid #d84315;"},

    7: {"grp": "Bewegung", "icon": "🏃", "name": "Start", "style": "background-color: #ffcdd2; color: #b71c1c; border: 1px solid #ef9a9a;"},
    8: {"grp": "Bewegung", "icon": "🏃", "name": "Ende",  "style": "background-color: #c62828; color: white; border: 1px solid #c62828;"},

    9:  {"grp": "Transition", "icon": "⏳", "name": "Start", "style": "background-color: #bbdefb; color: #0d47a1; border: 1px solid #90caf9;"},
    10: {"grp": "Transition", "icon": "⏳", "name": "Ende",  "style": "background-color: #1565c0; color: white; border: 1px solid #1565c0;"}
}

ALL_CATEGORIES = list(set(t["grp"] for t in TRIGGER_DEFS.values()))

# --------------------------
# HILFSFUNKTIONEN
# --------------------------

def load_data(uploaded_file):
    try:
        df = pd.read_csv(uploaded_file)
        df.columns = df.columns.str.strip()
        if 'TRIG' not in df.columns:
            return None, "❌ Spalte 'TRIG' fehlt."
        return df, None
    except Exception as e:
        return None, f"❌ Fehler: {e}"

def analyze_triggers(df, selected_ids):
    mask = df['TRIG'].isin(selected_ids)
    filtered = df[mask][['TRIG']].copy()
    
    analysis_data = []
    flow_sequence = []
    counts = {uid: 0 for uid in selected_ids}
    
    for idx, row in filtered.iterrows():
        val = int(row['TRIG'])
        counts[val] += 1
        
        info = TRIGGER_DEFS[val]
        desc = f"{info['icon']} {info['name']}"
        
        flow_sequence.append(str(val))
        
        analysis_data.append({
            "Zeile": idx,
            "Trigger": val,
            "Gruppe": info['grp'],
            "Beschreibung": desc,
            "Style": info['style']
        })
        
    return pd.DataFrame(analysis_data), flow_sequence, counts

# --------------------------
# UI & HAUPTPROGRAMM
# --------------------------

st.title("Trigger-Analyse")
st.markdown("Lade deine Playlists hoch und analysiere präzise die Abfolge der Trigger.")

# --- SIDEBAR EINSTELLUNGEN ---
with st.sidebar:
    st.header("⚙️ Einstellungen")
    st.write("Wähle die Trigger-Typen, die angezeigt werden sollen:")
    
    selected_cats = []
    defaults = ["Song", "Transition"]
    
    for cat in sorted(ALL_CATEGORIES):
        is_checked = st.checkbox(cat, value=(cat in defaults))
        if is_checked:
            selected_cats.append(cat)
    
    active_trigger_ids = [k for k, v in TRIGGER_DEFS.items() if v["grp"] in selected_cats]
    
    st.markdown("---")
    st.caption(f"Aktive Trigger-IDs: {active_trigger_ids}")

# --- UPLOAD BEREICH ---
col1, col2, col3 = st.columns(3)
file_yt = col1.file_uploader("YouTube (CSV)", type="csv")
file_a1 = col2.file_uploader("Audio 1 (CSV)", type="csv")
file_a2 = col3.file_uploader("Audio 2 (CSV)", type="csv")

files = [("YouTube", file_yt), ("Audio 1", file_a1), ("Audio 2", file_a2)]
loaded = [f for f in files if f[1] is not None]

st.markdown("---")

# --- ANALYSE BUTTON ---
if loaded:
    if st.button("Analyse starten", type="primary", use_container_width=True):
        st.header("📊 Ergebnisse")
        tabs = st.tabs([n for n, _ in loaded])
        
        for i, (name, file_obj) in enumerate(loaded):
            with tabs[i]:
                file_obj.seek(0)
                df, err = load_data(file_obj)
                
                if err:
                    st.error(err)
                else:
                    if not active_trigger_ids:
                        st.warning("⚠️ Bitte wähle links in den Einstellungen mindestens eine Trigger-Kategorie aus.")
                    else:
                        res_df, flow, counts = analyze_triggers(df, active_trigger_ids)
                        
                        if res_df.empty:
                            st.info(f"Keine Trigger der gewählten Kategorien in '{name}' gefunden.")
                        else:
                            st.subheader("Zählung")
                            
                            if len(selected_cats) > 0:
                                cols = st.columns(len(selected_cats))
                                for idx, cat in enumerate(selected_cats):
                                    with cols[idx]:
                                        cat_ids = [k for k, v in TRIGGER_DEFS.items() if v["grp"] == cat]
                                        
                                        group_icon = ""
                                        if cat_ids:
                                            group_icon = TRIGGER_DEFS[cat_ids[0]]['icon']

                                        st.markdown(f"**{cat} {group_icon}**")
                                        
                                        for cid in cat_ids:
                                            if cid in counts:
                                                entry = TRIGGER_DEFS[cid]
                                                st.write(f" {entry['name']}: **{counts[cid]}**")
                            
                            st.subheader("Sequenz-Fluss")
                            st.caption("Chronologische Reihenfolge der Trigger-IDs:")
                            st.code(" ➔ ".join(flow), language="text")
                            
                            st.subheader("📋 Detail-Liste")
                            
                            st.dataframe(
                                res_df.style.apply(lambda x: [x['Style']] * len(x), axis=1),
                                width="stretch", 
                                height=600,
                                column_config={
                                    "Zeile": st.column_config.NumberColumn(format="%d"),
                                    "Trigger": st.column_config.NumberColumn(format="%d"),
                                    "Style": None
                                }
                            )
else:
    st.info("Bitte CSV-Dateien hochladen, um zu beginnen.")