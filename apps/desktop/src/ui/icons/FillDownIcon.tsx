import { Icon, type IconProps } from "./Icon";

export function FillDownIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={4} y={3} width={8} height={8} rx={1} />
      <path d="M8 11v3" />
      <polyline points="6.5 13 8 14.5 9.5 13" />
    </Icon>
  );
}

