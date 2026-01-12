import { Icon, type IconProps } from "./Icon";

export function PageBreakIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={4} y={3} width={8} height={10} rx={1} />
      <path d="M5 8h6" strokeDasharray="2 1" />
    </Icon>
  );
}

