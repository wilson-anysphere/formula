import { Icon, type IconProps } from "./Icon";

export function AlignLeftIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M3 4h10" />
      <path d="M3 8h7" />
      <path d="M3 12h9" />
    </Icon>
  );
}
