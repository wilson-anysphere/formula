import { Icon, type IconProps } from "./Icon";

export function StopIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={5} y={5} width={6} height={6} rx={1} fill="currentColor" stroke="none" />
    </Icon>
  );
}

