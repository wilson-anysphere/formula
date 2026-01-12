import { Icon, type IconProps } from "./Icon";

export function SaveIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 2h7l2 2v10H4z" />
      <rect x={5.5} y={2.5} width={4.5} height={3} rx={0.5} />
      <rect x={6} y={9} width={4} height={4} rx={0.5} />
    </Icon>
  );
}

