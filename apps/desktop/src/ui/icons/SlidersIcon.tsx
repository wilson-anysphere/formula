import { Icon, type IconProps } from "./Icon";

export function SlidersIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M5 3v10" />
      <path d="M8 3v10" />
      <path d="M11 3v10" />
      <circle cx={5} cy={6} r={1} fill="currentColor" stroke="none" />
      <circle cx={8} cy={10} r={1} fill="currentColor" stroke="none" />
      <circle cx={11} cy={5} r={1} fill="currentColor" stroke="none" />
    </Icon>
  );
}

