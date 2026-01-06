import React from 'react';
import { Streamlit } from "streamlit-component-lib";
import { DndContext, closestCenter, KeyboardSensor, PointerSensor, useSensor, useSensors } from "@dnd-kit/core";
import { arrayMove, SortableContext, sortableKeyboardCoordinates, verticalListSortingStrategy, useSortable } from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import "./App.css";

function SortableAudioItem(props) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({ id: props.id });
  const { item, index, onBekanntChange, onGefallenChange, onEntspanntChange } = props;
  
  const style = { transform: CSS.Transform.toString(transform), transition };
  const itemClassName = `list-item ${isDragging ? "dragging" : ""}`;
  
  const isBaseline = item.id === 'baseline';

  const titleToDisplay = item.display_name ? item.display_name : item.name;

  return (
    <div ref={setNodeRef} style={style} className={itemClassName}>
      <div className="drag-handle" {...attributes} {...listeners}>⠿</div>
      <div className="item-rank">{index + 1}.</div>
      {item.image_data_url ? (
        <img src={item.image_data_url} alt={item.name} className="item-thumbnail" />
      ) : (
        <div className="item-thumbnail-placeholder">?</div>
      )}
      <div className="item-details">
        {/* Hier nutzen wir die neue Variable */}
        <div className="item-title">{titleToDisplay}</div>
        {!isBaseline && (
          item.audio_data_url ? (
              <audio controls src={item.audio_data_url} style={{ width: '100%', marginTop: '10px' }}>
                  Your browser does not support the audio element.
              </audio>
          ) : (
              <p style={{color: '#ff4b4b', fontSize: '0.9em'}}>Audio file not found.</p>
          )
        )}
      </div>
      
      {/* Controls Container */}
      <div className="item-controls">
        
        <div className="control-group">
            <input 
                type="checkbox" 
                id={`entspannt-${item.id}`} 
                checked={item.entspannt || false}
                onChange={(e) => onEntspanntChange(item.id, e.target.checked)}
                onClick={(e) => e.stopPropagation()} 
            />
            <label htmlFor={`entspannt-${item.id}`}>Entspannt</label>
        </div>

        {!isBaseline && (
        <>
          <div className="control-group">
              <input 
                  type="checkbox" 
                  id={`bekannt-${item.id}`} 
                  checked={item.bekannt}
                  onChange={(e) => onBekanntChange(item.id, e.target.checked)}
                  onClick={(e) => e.stopPropagation()}
              />
              <label htmlFor={`bekannt-${item.id}`}>Bekannt</label>
          </div>
          <div className="control-group">
              <label htmlFor={`gefallen-${item.id}`}>Gefallen?</label>
              <select
                  id={`gefallen-${item.id}`}
                  value={item.gefallen}
                  onChange={(e) => onGefallenChange(item.id, e.target.value)}
                  onClick={(e) => e.stopPropagation()}
              >
                  <option value="Neutral">Neutral</option>
                  <option value="Ja">Ja</option>
                  <option value="Nein">Nein</option>
              </select>
          </div>
        </>
        )}
      </div>
    </div>
  );
}

// Main component logic
function AudioRanker({ items }) {
  const sensors = useSensors(useSensor(PointerSensor), useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates }));

  const updateItemInStreamlit = (itemId, updateCallback) => {
    const updatedItems = items.map(item => 
      item.id === itemId ? updateCallback(item) : item
    );
    Streamlit.setComponentValue(updatedItems);
  };

  const handleBekanntChange = (itemId, newValue) => {
    updateItemInStreamlit(itemId, item => ({ ...item, bekannt: newValue }));
  };

  const handleGefallenChange = (itemId, newValue) => {
    updateItemInStreamlit(itemId, item => ({ ...item, gefallen: newValue }));
  };
  
  const handleEntspanntChange = (itemId, newValue) => {
    updateItemInStreamlit(itemId, item => ({ ...item, entspannt: newValue }));
  };

  const handleDragEnd = (event) => {
    const { active, over } = event;
    if (over && active.id !== over.id) {
      const oldIndex = items.findIndex(item => item.id === active.id);
      const newIndex = items.findIndex(item => item.id === over.id);
      const reorderedItems = arrayMove(items, oldIndex, newIndex);
      Streamlit.setComponentValue(reorderedItems);
    }
  };

  if (!items || items.length === 0) return null;
  
  return (
    <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={handleDragEnd}>
      <SortableContext items={items} strategy={verticalListSortingStrategy}>
        <div className="list-container">
          {items.map((item, index) => (
            <SortableAudioItem
              key={item.id} 
              id={item.id} 
              item={item} 
              index={index}
              onBekanntChange={handleBekanntChange}
              onGefallenChange={handleGefallenChange}
              onEntspanntChange={handleEntspanntChange} 
            />
          ))}
        </div>
      </SortableContext>
    </DndContext>
  );
}

export default AudioRanker;