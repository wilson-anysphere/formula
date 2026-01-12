import { Icon, type IconProps } from "./Icon";

export function ArrowLeftIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M12 8H4" />
      <polyline points="5.5 6.5 4 8 5.5 9.5" />
    </Icon>
  );
}

