import { Icon, type IconProps } from "./Icon";

export function ConditionalFormattingIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={3} width={10} height={10} rx={1} />
      <path d="M6 11V7" />
      <path d="M8 11V5.5" />
      <path d="M10 11V8.5" />
    </Icon>
  );
}

