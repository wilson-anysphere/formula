import { Icon, type IconProps } from "./Icon";

export function OrganizeSheetsIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={4} y={3} width={8} height={10} rx={1} />
      <path d="M3 5V4a1 1 0 0 1 1-1h6" />
      <path d="M5 8h6" />
      <path d="M5 10h4" />
    </Icon>
  );
}

