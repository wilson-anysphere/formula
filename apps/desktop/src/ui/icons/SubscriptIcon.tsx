import { Icon, type IconProps } from "./Icon";

export function SubscriptIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 4l6 8" />
      <path d="M10 4l-6 8" />
      <path d="M11 12.5h3" />
      <path d="M11 12.5c0-1 3-1 3-2.2 0-.7-.6-1.1-1.4-1.1-.7 0-1.2.2-1.6.6" />
    </Icon>
  );
}

