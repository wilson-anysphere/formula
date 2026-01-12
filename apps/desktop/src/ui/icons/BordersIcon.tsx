import { Icon, type IconProps } from "./Icon";

export function BordersIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={3} width={10} height={10} />
      <path d="M8 3v10" />
      <path d="M3 8h10" />
    </Icon>
  );
}
