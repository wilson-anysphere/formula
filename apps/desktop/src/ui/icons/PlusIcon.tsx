import { Icon, type IconProps } from "./Icon";

export function PlusIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M8 3v10" />
      <path d="M3 8h10" />
    </Icon>
  );
}

