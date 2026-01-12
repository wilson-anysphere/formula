import { Icon, type IconProps } from "./Icon";

export function ArrowRightIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 8h8" />
      <polyline points="10.5 6.5 12 8 10.5 9.5" />
    </Icon>
  );
}

