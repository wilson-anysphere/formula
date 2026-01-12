import { Icon, type IconProps } from "./Icon";

export function BringForwardIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={4} y={6} width={7} height={7} rx={1} />
      <rect x={3} y={4} width={7} height={7} rx={1} />
      <path d="M12.5 13V7" />
      <polyline points="11.5 8.5 12.5 7 13.5 8.5" />
    </Icon>
  );
}

