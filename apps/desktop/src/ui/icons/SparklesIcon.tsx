import { Icon, type IconProps } from "./Icon";

export function SparklesIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M8 3l1 2.5L11.5 6.5 9 7.5 8 10 7 7.5 4.5 6.5 7 5.5z" />
      <path d="M12 10.5l.5 1.3 1.3.5-1.3.5-.5 1.3-.5-1.3-1.3-.5 1.3-.5z" />
    </Icon>
  );
}

