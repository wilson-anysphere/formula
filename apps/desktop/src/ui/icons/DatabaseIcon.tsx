import { Icon, type IconProps } from "./Icon";

export function DatabaseIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <ellipse cx={8} cy={4.5} rx={4} ry={1.5} />
      <path d="M4 4.5v7c0 .8 1.8 1.5 4 1.5s4-.7 4-1.5v-7" />
      <path d="M4 8c0 .8 1.8 1.5 4 1.5s4-.7 4-1.5" />
    </Icon>
  );
}

