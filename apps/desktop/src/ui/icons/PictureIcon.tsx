import { Icon, type IconProps } from "./Icon";

export function PictureIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={3} width={10} height={10} rx={1} />
      <circle cx={6} cy={6} r={1} fill="currentColor" stroke="none" />
      <path d="M4 11l3-3 2 2 2-2 2 3" />
    </Icon>
  );
}

