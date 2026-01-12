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
      style={{ left: props.x, top: props.y }}
    >
      {props.preview}
    </div>
  );
}
