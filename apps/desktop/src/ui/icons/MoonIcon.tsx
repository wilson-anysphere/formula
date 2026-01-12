import { Icon, type IconProps } from "./Icon";

export function MoonIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M10.9 2.7a5.5 5.5 0 1 0 2.4 10.2 4.7 4.7 0 1 1-2.4-10.2z" />
    </Icon>
  );
}

