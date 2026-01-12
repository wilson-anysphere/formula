import { Icon, type IconProps } from "./Icon";

export function FeedbackIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 4h8a1 1 0 0 1 1 1v5a1 1 0 0 1-1 1H8l-3 3v-3H4a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1z" />
      <path d="M6 6.5h5" />
      <path d="M6 8.5h4" />
    </Icon>
  );
}

