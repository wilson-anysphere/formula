import { Icon, type IconProps } from "./Icon";

export function FunctionIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M7 4H4" />
      <path d="M5.5 4v8" />
      <path d="M4 7h3" />
      <path d="M9.5 6.5l3 3" />
      <path d="M12.5 6.5l-3 3" />
    </Icon>
  );
}

