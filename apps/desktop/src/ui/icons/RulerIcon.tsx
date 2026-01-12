import { Icon, type IconProps } from "./Icon";

export function RulerIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={5} width={10} height={6} rx={1} />
      <path d="M5 5v2" />
      <path d="M7 5v1.5" />
      <path d="M9 5v2" />
      <path d="M11 5v1.5" />
    </Icon>
  );
}

