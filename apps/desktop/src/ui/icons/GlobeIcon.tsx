import { Icon, type IconProps } from "./Icon";

export function GlobeIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <circle cx={8} cy={8} r={5} />
      <path d="M3 8h10" />
      <path d="M8 3v10" />
      <path d="M4.8 5.5c1 0.9 2.3 1.5 3.2 1.5s2.2-.6 3.2-1.5" />
      <path d="M4.8 10.5c1-0.9 2.3-1.5 3.2-1.5s2.2.6 3.2 1.5" />
    </Icon>
  );
}

