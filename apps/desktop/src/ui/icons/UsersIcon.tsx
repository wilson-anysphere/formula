import { Icon, type IconProps } from "./Icon";

export function UsersIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <circle cx={6} cy={6.5} r={1.7} />
      <circle cx={11} cy={7} r={1.3} />
      <path d="M2.8 13a3.2 3.2 0 0 1 6.4 0" />
      <path d="M9.3 13a2.7 2.7 0 0 1 5.4 0" />
    </Icon>
  );
}

