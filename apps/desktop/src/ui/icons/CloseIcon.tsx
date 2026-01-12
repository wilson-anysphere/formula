import { Icon, type IconProps } from "./Icon";

export function CloseIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 4l8 8" />
      <path d="M12 4l-8 8" />
    </Icon>
  );
}

