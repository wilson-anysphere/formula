import { Icon, type IconProps } from "./Icon";

export function CommaIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <ellipse cx={4.5} cy={7.5} rx={2} ry={3} />
      <ellipse cx={11.5} cy={7.5} rx={2} ry={3} />
      <circle cx={8.3} cy={9.4} r={0.75} fill="currentColor" stroke="none" />
      <path d="M8.3 10.2Q8.3 12 6.7 13.2" />
    </Icon>
  );
}
