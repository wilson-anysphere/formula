import { Icon, type IconProps } from "./Icon";

export function PenIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M8 2.5l3 3-3 8-3-8z" />
      <circle cx={8} cy={6.2} r={0.7} fill="currentColor" stroke="none" />
      <path d="M6 13h4" />
    </Icon>
  );
}

