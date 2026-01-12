import { Icon, type IconProps } from "./Icon";

export function InsertCellsIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={3} width={10} height={10} rx={1} />
      <path d="M8 6v4" />
      <path d="M6 8h4" />
    </Icon>
  );
}

