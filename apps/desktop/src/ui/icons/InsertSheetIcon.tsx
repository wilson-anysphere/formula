import { Icon, type IconProps } from "./Icon";

export function InsertSheetIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 2h6l3 3v9H4z" />
      <path d="M10 2v3h3" />
      <path d="M11 10.5v3" />
      <path d="M9.5 12h3" />
    </Icon>
  );
}

