import { Icon, type IconProps } from "./Icon";

export function PercentIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <circle cx={5} cy={5} r={1.5} />
      <circle cx={11} cy={11} r={1.5} />
      <path d="M12 4L4 12" />
    </Icon>
  );
}
