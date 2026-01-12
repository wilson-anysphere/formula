import { Icon, type IconProps } from "./Icon";

export function BookIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={4} width={5} height={9} rx={1} />
      <rect x={8} y={4} width={5} height={9} rx={1} />
      <path d="M8 4v9" />
    </Icon>
  );
}

