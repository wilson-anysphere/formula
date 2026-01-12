import { Icon, type IconProps } from "./Icon";

export function BoldIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M6 3v10" />
      <path d="M6 3h4a2.5 2.5 0 0 1 0 5H6" />
      <path d="M6 8h4.5a2.5 2.5 0 0 1 0 5H6" />
    </Icon>
  );
}
