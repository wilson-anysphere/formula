import { Icon, type IconProps } from "./Icon";

export function FontSizeIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 12L7 4l3 8" />
      <path d="M5.3 9h3.4" />
      <path d="M12 4v8" />
      <polyline points="11 5.5 12 4 13 5.5" />
      <polyline points="11 10.5 12 12 13 10.5" />
    </Icon>
  );
}

