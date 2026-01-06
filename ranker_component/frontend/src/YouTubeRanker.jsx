import React, { useEffect } from "react";
import { Streamlit } from "streamlit-component-lib";
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
} from "@dnd-kit/core";
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  verticalListSortingStrategy,
  useSortable,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import "./App.css";

// Sortable Item Component
function SortableItem(props) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({ id: props.id });
  // NEU: onEntspanntChange Prop
  const { item, index, onEntspanntChange } = props;

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
  };

  const itemClassName = `list-item ${isDragging ? "dragging" : ""}`;

  return (
    <div ref={setNodeRef} style={style} className={itemClassName}>
       <div className="drag-handle" {...attributes} {...listeners}>⠿</div>
      <div className="item-rank">{index + 1}.</div>
      <img src={item.thumbnail} alt={item.title} className="item-thumbnail" />
      <div className="item-details">
        <div className="item-title">{item.title}</div>
        {item.length && (
          <div className="item-length">Dauer: {item.length}</div>
        )}
      </div>

      {/* NEU: Controls Section für Checkbox */}
      <div className="item-controls" style={{ marginLeft: 'auto', paddingLeft: '10px' }}>
          <div className="control-group">
              <input 
                  type="checkbox" 
                  id={`entspannt-yt-${item.id}`} 
                  checked={item.entspannt || false}
                  onChange={(e) => onEntspanntChange(item.id, e.target.checked)}
                  onClick={(e) => e.stopPropagation()} // Verhindert Drag beim Klicken
              />
              <label htmlFor={`entspannt-yt-${item.id}`}>Entspannt</label>
          </div>
      </div>
    </div>
  );
}


// Main YouTube Ranker Component
function YouTubeRanker({ items }) {
  const sensors = useSensors(
    useSensor(PointerSensor),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  // NEU: Funktion zum Aktualisieren des Status (Checkbox) ohne Sortierung
  const updateItemInStreamlit = (itemId, updateCallback) => {
    const updatedItems = items.map(item => 
      item.id === itemId ? updateCallback(item) : item
    );
    // WICHTIG: Sende jetzt Items (Objekte), damit die Checkbox-Info ankommt.
    Streamlit.setComponentValue(updatedItems);
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
      
      // WICHTIG: Hier wurde geändert von reorderedIds zu reorderedItems (Objekte)
      // Damit wird der Checkbox Status nicht gelöscht beim Sortieren.
      Streamlit.setComponentValue(reorderedItems);
    }
  };

  useEffect(() => {
    Streamlit.setFrameHeight();
  });

  if (!items || items.length === 0) {
    return null;
  }
  
  return (
    <DndContext
      sensors={sensors}
      collisionDetection={closestCenter}
      onDragEnd={handleDragEnd}
    >
      <SortableContext items={items} strategy={verticalListSortingStrategy}>
        <div className="list-container">
          {items.map((item, index) => (
            <SortableItem 
                key={item.id} 
                id={item.id} 
                item={item} 
                index={index} 
                onEntspanntChange={handleEntspanntChange} // Prop übergeben
            />
          ))}
        </div>
      </SortableContext>
    </DndContext>
  );
}

export default YouTubeRanker;