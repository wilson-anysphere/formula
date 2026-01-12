import { Icon, type IconProps } from "./Icon";

export function WarningIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M8 2.5l6 11H2z" />
      <path d="M8 6v3.5" />
      <circle cx={8} cy={11.3} r={0.7} fill="currentColor" stroke="none" />
    </Icon>
  );
}

