import { Icon, type IconProps } from "./Icon";

export function SideBySideIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={2.5} y={4} width={5.5} height={9} rx={1} />
      <rect x={8} y={4} width={5.5} height={9} rx={1} />
    </Icon>
  );
}

