import { Icon, type IconProps } from "./Icon";

export function SnowflakeIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M8 2.5v11" />
      <path d="M4.5 4.5l7 7" />
      <path d="M11.5 4.5l-7 7" />
      <path d="M5.5 8h5" />
    </Icon>
  );
}

