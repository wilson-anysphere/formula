import { Icon, type IconProps } from "./Icon";

export function HelpIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <circle cx={8} cy={8} r={5.5} />
      <path d="M6.6 6.6a1.6 1.6 0 0 1 3 0c0 1.1-1.2 1.3-1.6 2.1-.1.2-.1.4-.1.8" />
      <circle cx={8} cy={11.2} r={0.7} fill="currentColor" stroke="none" />
    </Icon>
  );
}

