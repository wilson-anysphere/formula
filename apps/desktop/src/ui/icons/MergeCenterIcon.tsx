import { Icon, type IconProps } from "./Icon";

export function MergeCenterIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={4} width={10} height={8} />
      <path d="M8 4v2" />
      <path d="M8 10v2" />
      <path d="M6 6h4" />
      <path d="M4.5 8H7" />
      <polyline points="6 7 7 8 6 9" />
      <path d="M11.5 8H9" />
      <polyline points="10 7 9 8 10 9" />
    </Icon>
  );
}
