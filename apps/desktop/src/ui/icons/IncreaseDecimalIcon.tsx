import { Icon, type IconProps } from "./Icon";

export function IncreaseDecimalIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M7 4h6" />
      <polyline points="11.5 2.5 13 4 11.5 5.5" />

      <ellipse cx={4.5} cy={11.5} rx={2} ry={2.6} />
      <circle cx={7.6} cy={13.1} r={0.7} fill="currentColor" stroke="none" />
      <ellipse cx={11.2} cy={11.5} rx={2} ry={2.6} />
    </Icon>
  );
}
