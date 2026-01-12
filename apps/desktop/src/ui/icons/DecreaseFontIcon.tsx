import { Icon, type IconProps } from "./Icon";

export function DecreaseFontIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M3.5 13L6.5 5l3 8" />
      <path d="M4.8 10h3.4" />
      <path d="M11 6.5h3" />
    </Icon>
  );
}

