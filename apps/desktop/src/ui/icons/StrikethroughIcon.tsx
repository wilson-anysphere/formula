import { Icon, type IconProps } from "./Icon";

export function StrikethroughIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M3 8h10" />
      <path d="M11 5c0-1.1-1.3-2-3-2S5 3.9 5 5c0 1.6 2 2 3 2" />
      <path d="M5 11c0 1.1 1.3 2 3 2s3-.9 3-2c0-1.6-2-2-3-2" />
    </Icon>
  );
}

