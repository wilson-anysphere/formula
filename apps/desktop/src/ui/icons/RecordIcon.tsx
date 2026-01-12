import { Icon, type IconProps } from "./Icon";

export function RecordIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <circle cx={8} cy={8} r={3} fill="currentColor" stroke="none" />
    </Icon>
  );
}

