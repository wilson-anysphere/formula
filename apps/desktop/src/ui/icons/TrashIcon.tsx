import { Icon, type IconProps } from "./Icon";

export function TrashIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M5 5h6" />
      <path d="M6.5 5V3.5h3V5" />
      <rect x={5.5} y={5.5} width={5} height={8} rx={1} />
      <path d="M7 7.5v4" />
      <path d="M9 7.5v4" />
    </Icon>
  );
}

