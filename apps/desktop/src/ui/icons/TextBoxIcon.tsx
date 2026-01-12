import { Icon, type IconProps } from "./Icon";

export function TextBoxIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={3} width={10} height={10} rx={1} />
      <path d="M5 5h6" />
      <path d="M8 5v6" />
    </Icon>
  );
}

