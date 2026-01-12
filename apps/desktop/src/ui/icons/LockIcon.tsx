import { Icon, type IconProps } from "./Icon";

export function LockIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M5.5 7V5.5a2.5 2.5 0 0 1 5 0V7" />
      <rect x={4.5} y={7} width={7} height={7} rx={1} />
      <circle cx={8} cy={10} r={0.7} fill="currentColor" stroke="none" />
      <path d="M8 10.8v1.8" />
    </Icon>
  );
}

