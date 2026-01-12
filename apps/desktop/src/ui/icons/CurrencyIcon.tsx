import { Icon, type IconProps } from "./Icon";

export function CurrencyIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M8 2.5v11" />
      <path d="M10.5 4H7.5a2 2 0 0 0 0 4H8.5a2 2 0 0 1 0 4H5.5" />
    </Icon>
  );
}
