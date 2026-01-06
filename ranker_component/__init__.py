import os
import streamlit.components.v1 as components
import base64

_RELEASE = False 

if not _RELEASE:
    _ranker_component = components.declare_component("ranker_component", url="http://localhost:5173")
else:
    parent_dir = os.path.dirname(os.path.abspath(__file__))
    build_dir = os.path.join(parent_dir, "frontend/dist")
    _ranker_component = components.declare_component("ranker_component", path=build_dir)

def get_base64_data_url(file_path, mime_type):
    """Reads a file and returns it as a base64 data URL."""
    try:
        with open(file_path, "rb") as f:
            file_bytes = f.read()
        base64_encoded_data = base64.b64encode(file_bytes).decode('utf-8')
        return f"data:{mime_type};base64,{base64_encoded_data}"
    except FileNotFoundError:
        return None

def youtube_ranker(items, key=None):
    """Renders the YouTube video ranker."""
    return _ranker_component(items=items, component_type='youtube', key=key, default=items)

def audio_ranker(items, key=None):
    """
    Renders the local audio file ranker.
    """
    processed_items = []
    
    static_dir = os.path.join(os.getcwd(), 'static')
    
    for item in items:
        image_full_path = os.path.join(static_dir, item['image_path'])
        audio_full_path = os.path.join(static_dir, item['audio_path'])

        image_data_url = get_base64_data_url(image_full_path, 'image/jpeg')
        audio_data_url = get_base64_data_url(audio_full_path, "audio/mpeg")

        processed_items.append({
            "id": item['id'],
            "name": item['name'],
            "display_name": item.get('display_name', None), 
            "image_data_url": image_data_url,
            "audio_data_url": audio_data_url,
            "bekannt": item.get('bekannt', False),
            "gefallen": item.get('gefallen', 'Neutral'),
            "entspannt": item.get('entspannt', False)
        })
        
    return _ranker_component(
        items=processed_items,
        component_type='audio',
        key=key,
        default=items 
    )