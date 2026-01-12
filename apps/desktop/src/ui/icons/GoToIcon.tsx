import { Icon, type IconProps } from "./Icon";

export function GoToIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 12L12 4" />
      <polyline points="9 4 12 4 12 7" />
      <circle cx={12} cy={4} r={1} fill="currentColor" stroke="none" />
    </Icon>
  );
}

