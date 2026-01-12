import { Icon, type IconProps } from "./Icon";

export function FormatAsTableIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={3} width={10} height={10} rx={1} />
      <path d="M3 6.5h10" />
      <path d="M6.5 3v10" />
      <path d="M10 3v10" />
    </Icon>
  );
}

