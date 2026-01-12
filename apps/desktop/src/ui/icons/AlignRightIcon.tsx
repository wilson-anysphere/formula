import { Icon, type IconProps } from "./Icon";

export function AlignRightIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M3 4h10" />
      <path d="M6 8h7" />
      <path d="M4 12h9" />
    </Icon>
  );
}
