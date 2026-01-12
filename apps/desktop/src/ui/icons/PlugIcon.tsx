import { Icon, type IconProps } from "./Icon";

export function PlugIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M6 2.5v3" />
      <path d="M10 2.5v3" />
      <path d="M5 5.5h6v3a3 3 0 0 1-6 0z" />
      <path d="M8 11.5V14" />
    </Icon>
  );
}

