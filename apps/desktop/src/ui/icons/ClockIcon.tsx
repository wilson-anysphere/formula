import { Icon, type IconProps } from "./Icon";

export function ClockIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <circle cx={8} cy={8} r={5} />
      <path d="M8 5v3l2 1.5" />
    </Icon>
  );
}

