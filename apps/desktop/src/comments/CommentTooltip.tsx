import React from "react";

export interface CommentTooltipProps {
  x: number;
  y: number;
  preview: string;
}

type CommentTooltipStyle = React.CSSProperties & {
  "--comment-tooltip-x": string;
  "--comment-tooltip-y": string;
};

export function CommentTooltip(props: CommentTooltipProps): React.ReactElement {
  const style: CommentTooltipStyle = {
    "--comment-tooltip-x": `${props.x}px`,
    "--comment-tooltip-y": `${props.y}px`,
  };

  return (
    <div className="comment-tooltip comment-tooltip--visible" style={style}>
      {props.preview}
    </div>
  );
}
