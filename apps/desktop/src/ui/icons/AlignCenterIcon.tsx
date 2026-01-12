import { Icon, type IconProps } from "./Icon";

export function AlignCenterIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M3 4h10" />
      <path d="M4 8h8" />
      <path d="M5 12h6" />
    </Icon>
  );
}
