import { Icon, type IconProps } from "./Icon";

export function AutoSumIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M12 3H5l4 5-4 5h7" />
    </Icon>
  );
}

