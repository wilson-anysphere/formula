import { Icon, type IconProps } from "./Icon";

export function PrintIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={5} y={2} width={6} height={4} rx={1} />
      <rect x={4} y={6} width={8} height={5} rx={1} />
      <rect x={5} y={9} width={6} height={5} rx={1} />
      <circle cx={11} cy={7.5} r={0.6} fill="currentColor" stroke="none" />
    </Icon>
  );
}

