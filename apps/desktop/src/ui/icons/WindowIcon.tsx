import { Icon, type IconProps } from "./Icon";

export function WindowIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={4} width={10} height={9} rx={1} />
      <path d="M3 6h10" />
      <circle cx={5} cy={5} r={0.5} fill="currentColor" stroke="none" />
      <circle cx={6.5} cy={5} r={0.5} fill="currentColor" stroke="none" />
    </Icon>
  );
}

