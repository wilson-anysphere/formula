import { Icon, type IconProps } from "./Icon";

export function FillColorIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M9.5 2.5l4 4-6 6-4-4z" />
      <path d="M8 4l4 4" />
      <rect x={3} y={13} width={10} height={2} rx={1} fill="currentColor" stroke="none" />
    </Icon>
  );
}
