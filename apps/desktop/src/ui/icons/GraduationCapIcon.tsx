import { Icon, type IconProps } from "./Icon";

export function GraduationCapIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M2.5 6.5L8 4l5.5 2.5L8 9z" />
      <path d="M4 7.2V10c0 1 1.8 2 4 2s4-1 4-2V7.2" />
      <path d="M13.5 6.5v3.5" />
      <circle cx={13.5} cy={10.5} r={0.6} fill="currentColor" stroke="none" />
    </Icon>
  );
}

