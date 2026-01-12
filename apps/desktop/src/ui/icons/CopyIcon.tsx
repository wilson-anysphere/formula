import { Icon, type IconProps } from "./Icon";

export function CopyIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={5} y={5} width={8} height={9} rx={1} />
      <rect x={3} y={3} width={8} height={9} rx={1} />
    </Icon>
  );
}

