import { Icon, type IconProps } from "./Icon";

export function PrintAreaIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 2h6l3 3v9H4z" />
      <path d="M10 2v3h3" />
      <rect x={5.3} y={7.3} width={5.4} height={4.4} rx={0.6} />
    </Icon>
  );
}

