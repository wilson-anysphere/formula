import { Icon, type IconProps } from "./Icon";

export function UnderlineIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M5 3v6a3 3 0 0 0 6 0V3" />
      <path d="M4 13h8" />
    </Icon>
  );
}
