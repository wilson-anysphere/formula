import { Icon, type IconProps } from "./Icon";

export function MoreFormatsIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <circle cx={5} cy={8} r={1} fill="currentColor" stroke="none" />
      <circle cx={8} cy={8} r={1} fill="currentColor" stroke="none" />
      <circle cx={11} cy={8} r={1} fill="currentColor" stroke="none" />
    </Icon>
  );
}

