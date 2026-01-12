import { Icon, type IconProps } from "./Icon";

export function CodeIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <polyline points="6 5 4 8 6 11" />
      <polyline points="10 5 12 8 10 11" />
      <path d="M8.8 4.5l-1.6 7" />
    </Icon>
  );
}

