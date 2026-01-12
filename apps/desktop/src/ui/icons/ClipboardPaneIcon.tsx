import { Icon, type IconProps } from "./Icon";

export function ClipboardPaneIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={4} y={3} width={8} height={11} rx={1.5} />
      <rect x={6} y={1.5} width={4} height={3} rx={1} />
      <path d="M6 6.5h4" />
      <path d="M6 8.5h4" />
      <path d="M6 10.5h3" />
    </Icon>
  );
}

