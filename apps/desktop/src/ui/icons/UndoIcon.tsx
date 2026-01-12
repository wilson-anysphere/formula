import { Icon, type IconProps } from "./Icon";

export function UndoIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M6 5H3v-3" />
      <path d="M3 5c1.2-1.5 3.1-2.5 5.2-2.5 3 0 5.3 2 5.3 5 0 3-2.4 5-5.3 5H6" />
    </Icon>
  );
}

