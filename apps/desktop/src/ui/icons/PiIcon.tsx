import { Icon, type IconProps } from "./Icon";

export function PiIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 5h8" />
      <path d="M6 5v8" />
      <path d="M10 5v8" />
    </Icon>
  );
}

