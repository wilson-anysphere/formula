import { Icon, type IconProps } from "./Icon";

export function LightningIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M9 2L4.5 9H8l-1 5L11.5 7H8z" />
    </Icon>
  );
}

