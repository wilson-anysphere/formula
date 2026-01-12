import { Icon, type IconProps } from "./Icon";

export function PhoneIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M6 3h4v10H6z" />
      <circle cx={8} cy={11.5} r={0.6} fill="currentColor" stroke="none" />
    </Icon>
  );
}

