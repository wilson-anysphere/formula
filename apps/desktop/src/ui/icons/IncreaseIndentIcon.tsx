import { Icon, type IconProps } from "./Icon";

export function IncreaseIndentIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M3 4h10" />
      <path d="M6 8h7" />
      <path d="M3 12h10" />
      <polyline points="3.5 6.5 5.5 8 3.5 9.5" />
    </Icon>
  );
}
