import { Icon, type IconProps } from "./Icon";

export function RefreshIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M12.5 8a4.5 4.5 0 1 1-1.1-2.9" />
      <polyline points="12.5 3.5 12.5 6 10 6" />
    </Icon>
  );
}

