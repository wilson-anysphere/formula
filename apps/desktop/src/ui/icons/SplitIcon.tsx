import { Icon, type IconProps } from "./Icon";

export function SplitIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={3} width={10} height={10} rx={1} />
      <path d="M8 3v10" />
      <path d="M3 8h10" />
    </Icon>
  );
}

