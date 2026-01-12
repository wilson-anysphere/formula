import { Icon, type IconProps } from "./Icon";

export function SendBackwardIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={4} y={6} width={7} height={7} rx={1} />
      <rect x={3} y={4} width={7} height={7} rx={1} />
      <path d="M12.5 7v6" />
      <polyline points="11.5 11.5 12.5 13 13.5 11.5" />
    </Icon>
  );
}

