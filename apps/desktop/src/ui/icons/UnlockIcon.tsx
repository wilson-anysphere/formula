import { Icon, type IconProps } from "./Icon";

export function UnlockIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M6 7V5.8a2.6 2.6 0 0 1 4.7-1.6" />
      <rect x={4.5} y={7} width={7} height={7} rx={1} />
      <circle cx={8} cy={10} r={0.7} fill="currentColor" stroke="none" />
      <path d="M8 10.8v1.8" />
    </Icon>
  );
}

