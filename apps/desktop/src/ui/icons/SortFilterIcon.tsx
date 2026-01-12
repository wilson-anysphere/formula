import { Icon, type IconProps } from "./Icon";

export function SortFilterIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 4h6" />
      <path d="M4 7h4" />
      <path d="M4 10h5" />
      <path d="M11 4h3l-1.5 2v4l-1-0.5V6z" />
    </Icon>
  );
}

