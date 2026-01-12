import { Icon, type IconProps } from "./Icon";

export function TagIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M3 8V4h4l6 6-4 4-6-6z" />
      <circle cx={5.5} cy={5.5} r={0.7} fill="currentColor" stroke="none" />
    </Icon>
  );
}

