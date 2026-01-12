import { Icon, type IconProps } from "./Icon";

export function CalculatorIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <rect x={4} y={2.5} width={8} height={11} rx={1} />
      <rect x={5.5} y={4} width={5} height={2} rx={0.5} />
      <circle cx={6.5} cy={8} r={0.6} fill="currentColor" stroke="none" />
      <circle cx={8} cy={8} r={0.6} fill="currentColor" stroke="none" />
      <circle cx={9.5} cy={8} r={0.6} fill="currentColor" stroke="none" />
      <circle cx={6.5} cy={10} r={0.6} fill="currentColor" stroke="none" />
      <circle cx={8} cy={10} r={0.6} fill="currentColor" stroke="none" />
      <circle cx={9.5} cy={10} r={0.6} fill="currentColor" stroke="none" />
      <circle cx={6.5} cy={12} r={0.6} fill="currentColor" stroke="none" />
      <circle cx={8} cy={12} r={0.6} fill="currentColor" stroke="none" />
      <circle cx={9.5} cy={12} r={0.6} fill="currentColor" stroke="none" />
    </Icon>
  );
}

