import { Icon, type IconProps } from "./Icon";

export function HashIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M6 3L5 13" />
      <path d="M11 3l-1 10" />
      <path d="M4 6h9" />
      <path d="M3.5 10h9" />
    </Icon>
  );
}

