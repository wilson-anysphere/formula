import { Icon, type IconProps } from "./Icon";

export function InsertColumnsIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={2} width={10} height={9} rx={1} />
      <path d="M6.5 2v9" />
      <path d="M9.5 2v9" />
      <path d="M6.5 13h3" />
      <path d="M8 11.5v3" />
    </Icon>
  );
}

