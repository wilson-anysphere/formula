import { Icon, type IconProps } from "./Icon";

export function ShieldIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M8 2.5l4 1.8v4.1c0 2.6-1.6 4.7-4 5.6-2.4-.9-4-3-4-5.6V4.3z" />
      <path d="M8 5v5" />
      <path d="M6.5 8h3" />
    </Icon>
  );
}

