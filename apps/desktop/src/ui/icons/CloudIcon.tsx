import { Icon, type IconProps } from "./Icon";

export function CloudIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M5.5 12.5H12a2.5 2.5 0 0 0 .3-5 3.5 3.5 0 0 0-6.8-1A2.5 2.5 0 0 0 5.5 12.5z" />
    </Icon>
  );
}

