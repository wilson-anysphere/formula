import { Icon, type IconProps } from "./Icon";

export function StarIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path
        d="M8 2.4l1.6 3.3 3.6.5-2.6 2.5.6 3.6L8 10.7l-3.2 1.6.6-3.6L2.8 6.2l3.6-.5z"
        fill="currentColor"
        stroke="none"
      />
    </Icon>
  );
}

