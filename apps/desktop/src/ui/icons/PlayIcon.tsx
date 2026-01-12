import { Icon, type IconProps } from "./Icon";

export function PlayIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M6 4l7 4-7 4z" fill="currentColor" stroke="none" />
    </Icon>
  );
}

