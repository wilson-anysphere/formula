import { Icon, type IconProps } from "./Icon";

export function SmartArtIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <circle cx={5} cy={5.5} r={1.3} fill="currentColor" stroke="none" />
      <circle cx={11} cy={5.5} r={1.3} fill="currentColor" stroke="none" />
      <circle cx={8} cy={11} r={1.3} fill="currentColor" stroke="none" />
      <path d="M6.6 5.5h2.8" />
      <path d="M5.8 6.8l1.6 3.2" />
      <path d="M10.2 6.8l-1.6 3.2" />
    </Icon>
  );
}

