import { Icon, type IconProps } from "./Icon";

export function FilterIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M3 4h10L9 8.5V13l-2-1V8.5L3 4z" />
    </Icon>
  );
}
