import { Icon, type IconProps } from "./Icon";

export function PencilIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M10.5 3.5l2 2-7 7H3.5v-2z" />
      <path d="M9 5l2 2" />
    </Icon>
  );
}

