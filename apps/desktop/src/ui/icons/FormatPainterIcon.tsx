import { Icon, type IconProps } from "./Icon";

export function FormatPainterIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={3} width={7} height={3} rx={1} />
      <path d="M10 4h3v4H9" />
      <path d="M9 8v5" />
      <path d="M8 13h2" />
    </Icon>
  );
}

