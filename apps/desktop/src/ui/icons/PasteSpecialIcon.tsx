import { Icon, type IconProps } from "./Icon";

export function PasteSpecialIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={4} y={3} width={8} height={11} rx={1.5} />
      <rect x={6} y={1.5} width={4} height={3} rx={1} />
      <path d="M11 11.5h2.5" />
      <path d="M12.25 10.25v2.5" />
    </Icon>
  );
}

