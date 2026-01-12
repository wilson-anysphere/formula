import { Icon, type IconProps } from "./Icon";

export function PinIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M8 14s4-3.5 4-6.5A4 4 0 0 0 4 7.5C4 10.5 8 14 8 14z" />
      <circle cx={8} cy={7.5} r={1.3} />
    </Icon>
  );
}

