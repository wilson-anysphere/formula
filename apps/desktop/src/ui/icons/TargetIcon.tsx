import { Icon, type IconProps } from "./Icon";

export function TargetIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <circle cx={8} cy={8} r={5} />
      <circle cx={8} cy={8} r={2.5} />
      <circle cx={8} cy={8} r={0.8} fill="currentColor" stroke="none" />
    </Icon>
  );
}

