import { Icon, type IconProps } from "./Icon";

export function MinusIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 8h8" />
    </Icon>
  );
}

