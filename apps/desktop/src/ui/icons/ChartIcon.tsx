import { Icon, type IconProps } from "./Icon";

export function ChartIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M3 13V3" />
      <path d="M3 13h10" />
      <path d="M5 12V9" strokeLinecap="butt" />
      <path d="M8 12V7" strokeLinecap="butt" />
      <path d="M11 12V5" strokeLinecap="butt" />
    </Icon>
  );
}

