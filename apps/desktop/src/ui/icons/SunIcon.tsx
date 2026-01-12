import { Icon, type IconProps } from "./Icon";

export function SunIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <circle cx="8" cy="8" r="2.5" />
      <path d="M8 1.5v2" />
      <path d="M8 12.5v2" />
      <path d="M1.5 8h2" />
      <path d="M12.5 8h2" />
      <path d="M3.2 3.2l1.4 1.4" />
      <path d="M11.4 11.4l1.4 1.4" />
      <path d="M12.8 3.2l-1.4 1.4" />
      <path d="M4.6 11.4l-1.4 1.4" />
    </Icon>
  );
}

