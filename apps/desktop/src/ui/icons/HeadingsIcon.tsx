import { Icon, type IconProps } from "./Icon";

export function HeadingsIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={3} width={10} height={10} rx={1} />
      <path d="M3 5.5h10" strokeWidth={2} />
      <path d="M5.5 3v10" strokeWidth={2} />
      <path d="M5.5 9h7.5" />
      <path d="M9 5.5v7.5" />
    </Icon>
  );
}

