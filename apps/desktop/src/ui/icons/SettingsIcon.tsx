import { Icon, type IconProps } from "./Icon";

export function SettingsIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <circle cx={8} cy={8} r={2.5} />
      <path d="M8 2.5v1.5" />
      <path d="M8 12v1.5" />
      <path d="M2.5 8h1.5" />
      <path d="M12 8h1.5" />
      <path d="M4.2 4.2l1.1 1.1" />
      <path d="M10.7 10.7l1.1 1.1" />
      <path d="M11.8 4.2l-1.1 1.1" />
      <path d="M5.3 10.7l-1.1 1.1" />
    </Icon>
  );
}

