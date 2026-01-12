import { Icon, type IconProps } from "./Icon";

export function PuzzleIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M6 3h3a1.3 1.3 0 0 1 2.6 0H13v3a1.3 1.3 0 0 0 0 2.6V13h-3a1.3 1.3 0 0 0-2.6 0H6v-3a1.3 1.3 0 0 1 0-2.6z" />
    </Icon>
  );
}

