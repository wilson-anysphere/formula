import { Icon, type IconProps } from "./Icon";

export function UserIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <circle cx={8} cy={6} r={2} />
      <path d="M4 14a4 4 0 0 1 8 0" />
    </Icon>
  );
}

