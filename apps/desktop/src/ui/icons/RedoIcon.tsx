import { Icon, type IconProps } from "./Icon";

export function RedoIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M10 5h3v-3" />
      <path d="M13 5c-1.2-1.5-3.1-2.5-5.2-2.5-3 0-5.3 2-5.3 5 0 3 2.4 5 5.3 5H10" />
    </Icon>
  );
}

