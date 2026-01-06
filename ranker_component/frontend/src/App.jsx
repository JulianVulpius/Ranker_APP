import React, { useEffect } from "react";
import { Streamlit, withStreamlitConnection } from "streamlit-component-lib";

import YouTubeRanker from "./YouTubeRanker";
import AudioRanker from "./AudioRanker";

function App(props) {
  const { component_type, items } = props.args;

  useEffect(() => {
    Streamlit.setFrameHeight();
  });

  switch (component_type) {
    case "youtube":
      return <YouTubeRanker items={items} />;
    case "audio":
      return <AudioRanker items={items} />;
    default:
      return <div>Error: Unknown component type!</div>;
  }
}

export default withStreamlitConnection(App);