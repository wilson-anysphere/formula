import { Icon, type IconProps } from "./Icon";

export function FontColorIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4.5 13L8 3l3.5 10" />
      <path d="M6 9h4" />
      <rect x={3} y={14} width={10} height={2} rx={1} fill="currentColor" stroke="none" />
    </Icon>
  );
}
