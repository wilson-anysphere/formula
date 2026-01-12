import { Icon, type IconProps } from "./Icon";

export function ShareIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={7} width={7} height={6} rx={1} />
      <path d="M9 3h4v4" />
      <path d="M13 3L7 9" />
    </Icon>
  );
}

