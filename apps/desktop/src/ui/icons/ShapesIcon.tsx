import { Icon, type IconProps } from "./Icon";

export function ShapesIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={4} width={6} height={6} rx={1} />
      <circle cx={10.5} cy={10} r={3} />
    </Icon>
  );
}

