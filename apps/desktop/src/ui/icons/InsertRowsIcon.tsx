import { Icon, type IconProps } from "./Icon";

export function InsertRowsIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={2} y={3} width={9} height={10} rx={1} />
      <path d="M2 6.5h9" />
      <path d="M2 9.5h9" />
      <path d="M13 6.5v3" />
      <path d="M11.5 8h3" />
    </Icon>
  );
}

