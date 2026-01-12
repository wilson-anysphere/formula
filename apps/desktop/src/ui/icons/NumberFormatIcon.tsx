import { Icon, type IconProps } from "./Icon";

export function NumberFormatIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 5h3" />
      <path d="M5.5 5v6" />
      <path d="M4 11h3" />

      <path d="M10 5h2a1.5 1.5 0 0 1 0 3H10v3h3" />
    </Icon>
  );
}

