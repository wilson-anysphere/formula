import { Icon, type IconProps } from "./Icon";

export function CalendarIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={4} width={10} height={9} rx={1} />
      <path d="M5 3v2" />
      <path d="M11 3v2" />
      <path d="M3 6.5h10" />
      <path d="M5.5 8.5h2" />
      <path d="M8.5 8.5h2" />
      <path d="M5.5 11h2" />
    </Icon>
  );
}

