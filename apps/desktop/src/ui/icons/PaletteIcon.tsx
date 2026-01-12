import { Icon, type IconProps } from "./Icon";

export function PaletteIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M8 2.5c-3.3 0-6 2.3-6 5.2 0 2.5 2 4.3 4.6 4.3H7c.9 0 1.6.6 1.6 1.5 0 .6-.3 1.1-.8 1.3 3-.2 6.2-2.2 6.2-6.5 0-3.1-2.5-5.8-6-5.8z" />
      <circle cx={5} cy={7} r={0.7} fill="currentColor" stroke="none" />
      <circle cx={7} cy={5} r={0.7} fill="currentColor" stroke="none" />
      <circle cx={9.5} cy={5.2} r={0.7} fill="currentColor" stroke="none" />
    </Icon>
  );
}

