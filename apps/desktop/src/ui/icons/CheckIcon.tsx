import { Icon, type IconProps } from "./Icon";

export function CheckIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 8l2.5 2.5L12 5" />
    </Icon>
  );
}

