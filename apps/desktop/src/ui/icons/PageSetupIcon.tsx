import { Icon, type IconProps } from "./Icon";

export function PageSetupIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 2h6l3 3v9H4z" />
      <path d="M10 2v3h3" />
      <path d="M6 8h6" />
      <path d="M6 10h5" />
      <path d="M6 12h4" />
    </Icon>
  );
}

