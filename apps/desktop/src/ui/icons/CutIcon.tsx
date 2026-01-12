import { Icon, type IconProps } from "./Icon";

export function CutIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <circle cx={5} cy={5} r={1.5} />
      <circle cx={5} cy={11} r={1.5} />
      <path d="M6.2 6.2L7 8" />
      <path d="M6.2 9.8L7 8" />
      <path d="M7 8L13 3" />
      <path d="M7 8L13 13" />
    </Icon>
  );
}

