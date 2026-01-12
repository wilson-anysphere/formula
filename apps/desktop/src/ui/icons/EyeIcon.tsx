import { Icon, type IconProps } from "./Icon";

export function EyeIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M2 8c1.5-2.5 3.5-4 6-4s4.5 1.5 6 4c-1.5 2.5-3.5 4-6 4S3.5 10.5 2 8z" />
      <circle cx={8} cy={8} r={1.5} />
    </Icon>
  );
}

