import { Icon, type IconProps } from "./Icon";

export function MailIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={4} width={10} height={8} rx={1} />
      <path d="M3 5l5 4 5-4" />
      <path d="M3 12l4.5-3.5" />
      <path d="M13 12l-4.5-3.5" />
    </Icon>
  );
}

