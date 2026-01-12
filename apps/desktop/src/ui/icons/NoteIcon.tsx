import { Icon, type IconProps } from "./Icon";

export function NoteIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={4} y={3} width={8} height={10} rx={1} />
      <path d="M6 5h4" />
      <path d="M6 7h4" />
      <path d="M6 9h3" />
    </Icon>
  );
}

