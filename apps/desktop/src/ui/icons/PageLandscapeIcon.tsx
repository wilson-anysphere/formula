import { Icon, type IconProps } from "./Icon";

export function PageLandscapeIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={3} y={5} width={10} height={6} rx={1} />
    </Icon>
  );
}

