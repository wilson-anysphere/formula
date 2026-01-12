import { Icon, type IconProps } from "./Icon";

export function WrapTextIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M3 4h10" />
      <path d="M3 7h7" />
      <path d="M13 7v4H9" />
      <polyline points="10.5 9.5 9 11 10.5 12.5" />
      <path d="M3 12h10" />
    </Icon>
  );
}
