import type { CSSProperties } from "react";
import React from "react";

export interface CommentTooltipProps {
  x: number;
  y: number;
  preview: string;
}

const containerStyle: CSSProperties = {
  position: "absolute",
  maxWidth: 320,
  padding: "8px 10px",
  background: "rgba(17, 24, 39, 0.95)",
  color: "white",
  borderRadius: 8,
  fontSize: 12,
  lineHeight: 1.4,
  pointerEvents: "none",
  zIndex: 50,
  whiteSpace: "pre-wrap",
};

export function CommentTooltip(props: CommentTooltipProps): React.ReactElement {
  return (
    <div style={{ ...containerStyle, left: props.x, top: props.y }}>
      {props.preview}
    </div>
  );
}

