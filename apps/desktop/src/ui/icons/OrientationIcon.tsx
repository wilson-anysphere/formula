import { Icon, type IconProps } from "./Icon";

export function OrientationIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 13V3" />
      <path d="M7 12l5-5" />
      <path d="M11 6.5h2.5" />
      <path d="M12 9.5h1.5" />
    </Icon>
  );
}

