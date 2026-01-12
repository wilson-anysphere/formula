import { Icon, type IconProps } from "./Icon";

export function ExportIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={4} width={7} height={8} rx={1} />
      <path d="M8 8h5" />
      <polyline points="11.5 6.5 13 8 11.5 9.5" />
    </Icon>
  );
}

