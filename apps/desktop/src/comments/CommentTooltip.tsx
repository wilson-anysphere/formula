import React from "react";

export interface CommentTooltipProps {
  x: number;
  y: number;
  preview: string;
}

export function CommentTooltip(props: CommentTooltipProps): React.ReactElement {
  return (
    <div
      className="comment-tooltip comment-tooltip--visible"
      style={
        {
          "--comment-tooltip-x": `${props.x}px`,
          "--comment-tooltip-y": `${props.y}px`,
        } as React.CSSProperties
      }
    >
      {props.preview}
    </div>
  );
}
