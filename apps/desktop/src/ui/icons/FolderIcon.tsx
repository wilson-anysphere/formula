import { Icon, type IconProps } from "./Icon";

export function FolderIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M3 5V3.5h4l1 1.5h5v8H3z" />
    </Icon>
  );
}

