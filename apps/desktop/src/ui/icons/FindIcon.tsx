import { Icon, type IconProps } from "./Icon";

export function FindIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <circle cx={7} cy={7} r={3.5} />
      <path d="M10 10l3 3" />
    </Icon>
  );
}
