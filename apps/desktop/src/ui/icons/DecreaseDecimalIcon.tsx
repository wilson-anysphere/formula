import { Icon, type IconProps } from "./Icon";

export function DecreaseDecimalIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M13 4H7" />
      <polyline points="8.5 2.5 7 4 8.5 5.5" />

      <ellipse cx={4.5} cy={11.5} rx={2} ry={2.6} />
      <circle cx={7.6} cy={13.1} r={0.7} fill="currentColor" stroke="none" />
      <ellipse cx={11.2} cy={11.5} rx={2} ry={2.6} />
    </Icon>
  );
}
