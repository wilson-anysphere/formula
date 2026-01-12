import * as React from "react";

export type IconProps = React.SVGProps<SVGSVGElement> & {
  /**
   * Icon size in CSS pixels. Defaults to 16 for ribbon controls, but the icons
   * are designed to also look good at 20.
   */
  size?: number;
  /**
   * Optional accessible title. When omitted, the icon is marked as decorative
   * unless an `aria-label` / `aria-labelledby` is provided.
   */
  title?: string;
};

export function Icon({
  size = 16,
  strokeWidth = 1.5,
  title,
  children,
  width,
  height,
  ...svgProps
}: IconProps) {
  const titleId = React.useId();
  const ariaLabel = svgProps["aria-label"];
  const ariaLabelledBy = svgProps["aria-labelledby"];
  const ariaHidden = svgProps["aria-hidden"];

  const computedAriaLabelledBy = title ? titleId : ariaLabelledBy;
  const computedAriaHidden = ariaHidden ?? (title || ariaLabel || ariaLabelledBy ? undefined : true);

  return (
    <svg
      {...svgProps}
      width={width ?? size}
      height={height ?? size}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth={strokeWidth}
      strokeLinecap="round"
      strokeLinejoin="round"
      role="img"
      aria-hidden={computedAriaHidden}
      aria-labelledby={computedAriaLabelledBy}
    >
      {title ? <title id={titleId}>{title}</title> : null}
      {children}
    </svg>
  );
}
