import { Icon, type IconProps } from "./Icon";

export function SortIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M3 4h6" />
      <path d="M3 8h4" />
      <path d="M3 12h5" />

      <path d="M11.5 12V4" />
      <polyline points="10.5 5.5 11.5 4 12.5 5.5" />

      <path d="M13.5 4v8" />
      <polyline points="12.5 10.5 13.5 12 14.5 10.5" />
    </Icon>
  );
}
