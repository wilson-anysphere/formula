import { Icon, type IconProps } from "./Icon";

export function FileIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 2h6l3 3v9H4z" />
      <path d="M10 2v3h3" />
    </Icon>
  );
}

