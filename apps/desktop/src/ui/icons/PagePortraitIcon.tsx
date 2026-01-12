import { Icon, type IconProps } from "./Icon";

export function PagePortraitIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={5} y={3} width={6} height={10} rx={1} />
    </Icon>
  );
}

