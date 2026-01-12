import { Icon, type IconProps } from "./Icon";

export function ReplaceIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M3 5h8" />
      <polyline points="9.5 3.5 11 5 9.5 6.5" />
      <path d="M13 11H5" />
      <polyline points="6.5 9.5 5 11 6.5 12.5" />
    </Icon>
  );
}

