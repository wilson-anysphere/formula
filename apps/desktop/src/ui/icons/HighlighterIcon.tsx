import { Icon, type IconProps } from "./Icon";

export function HighlighterIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M9 3l4 4-5 5-4-4z" />
      <path d="M4 8l-1 1v2h2l1-1" />
      <path d="M3 13h10" />
    </Icon>
  );
}

