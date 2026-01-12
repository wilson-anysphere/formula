import { Icon, type IconProps } from "../../ui/icons/Icon";

type RibbonSvgIconProps = Omit<IconProps, "children">;

export function FileIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M5 1.5h4.5L13 5v9.5A1.5 1.5 0 0 1 11.5 16h-6A1.5 1.5 0 0 1 4 14.5V3A1.5 1.5 0 0 1 5.5 1.5" />
      <path d="M9.5 1.5V5H13" />
    </Icon>
  );
}

export function FilePlusIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M5 1.5h4.5L13 5v9.5A1.5 1.5 0 0 1 11.5 16h-6A1.5 1.5 0 0 1 4 14.5V3A1.5 1.5 0 0 1 5.5 1.5" />
      <path d="M9.5 1.5V5H13" />
      <path d="M8 8v4" />
      <path d="M6 10h4" />
    </Icon>
  );
}

export function FolderOpenIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M2.5 5.5A2 2 0 0 1 4.5 3.5h2l1.5 1.5h5A2 2 0 0 1 15 7v6.5A2 2 0 0 1 13 15.5H4.5a2 2 0 0 1-2-2V5.5z" />
      <path d="M2.5 7.5h13" />
    </Icon>
  );
}

export function CalendarIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <rect x={3} y={4} width={10} height={9} rx={1.5} />
      <path d="M3 6.5h10" />
      <path d="M5.5 2.5v3" />
      <path d="M10.5 2.5v3" />
    </Icon>
  );
}

export function SaveIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M3.5 2.5h9l1.5 1.5v10.5a1.5 1.5 0 0 1-1.5 1.5h-9A1.5 1.5 0 0 1 2 14.5V4A1.5 1.5 0 0 1 3.5 2.5" />
      <path d="M5 2.5V7h6V2.5" />
      <rect x={5} y={10} width={6} height={4} rx={1} />
    </Icon>
  );
}

export function EditIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M4 12l-.5 2.5L6 14l6.5-6.5-2-2L4 12z" />
      <path d="M9.5 5.5l2 2" />
    </Icon>
  );
}

export function PrintIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M5 5V2.5h6V5" />
      <rect x={4} y={10} width={8} height={5} rx={1} />
      <path d="M4 10H3.5A1.5 1.5 0 0 1 2 8.5V6.5A1.5 1.5 0 0 1 3.5 5h9A1.5 1.5 0 0 1 14 6.5v2A1.5 1.5 0 0 1 12.5 10H12" />
      <path d="M11.5 7.25h.01" />
    </Icon>
  );
}

export function SettingsIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M8 3.25l.8 1.38 1.6.34.45 1.55 1.33.92-.73 1.44.73 1.44-1.33.92-.45 1.55-1.6.34-.8 1.38-.8-1.38-1.6-.34-.45-1.55-1.33-.92.73-1.44-.73-1.44 1.33-.92.45-1.55 1.6-.34L8 3.25z" />
      <circle cx={8} cy={8} r={1.75} />
    </Icon>
  );
}

export function LockIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <rect x={4} y={7} width={8} height={8} rx={1.5} />
      <path d="M6 7V5.5A2 2 0 0 1 8 3.5a2 2 0 0 1 2 2V7" />
    </Icon>
  );
}

export function SearchIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <circle cx={7} cy={7} r={3} />
      <path d="M9.5 9.5L13.5 13.5" />
    </Icon>
  );
}

export function MenuIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M3 4h10" />
      <path d="M3 8h10" />
      <path d="M3 12h10" />
    </Icon>
  );
}

export function ClockIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <circle cx={8} cy={8} r={5.5} />
      <path d="M8 5.5V8.5" />
      <path d="M8 8.5L10 10" />
    </Icon>
  );
}

export function PinIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M6 2.5h4l-.5 3 2 2-3 1v7l-1-1-1 1v-7l-3-1 2-2-.5-3z" />
    </Icon>
  );
}

export function TagIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M3 8l5-5h5v5l-5 5L3 8z" />
      <circle cx={11} cy={5} r={0.75} />
    </Icon>
  );
}

export function UploadIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M4 13.5h8" />
      <path d="M8 12.5V4.5" />
      <path d="M5.5 7L8 4.5 10.5 7" />
    </Icon>
  );
}

export function RefreshIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M13 8a5 5 0 1 1-1.1-3.1" />
      <path d="M12.5 3.5v2.5H10" />
    </Icon>
  );
}

export function DownloadIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M4 13.5h8" />
      <path d="M8 3.5v8" />
      <path d="M5.5 9.5L8 12l2.5-2.5" />
    </Icon>
  );
}

export function LinkIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M6.5 9.5l-1 1a2.5 2.5 0 0 1-3.5 0 2.5 2.5 0 0 1 0-3.5l1-1" />
      <path d="M9.5 6.5l1-1a2.5 2.5 0 0 1 3.5 0 2.5 2.5 0 0 1 0 3.5l-1 1" />
      <path d="M6 10l4-4" />
    </Icon>
  );
}

export function MailIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <rect x={2.5} y={4.5} width={11} height={8} rx={1.5} />
      <path d="M3.5 5.5L8 9l4.5-3.5" />
    </Icon>
  );
}

export function GlobeIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <circle cx={8} cy={8} r={5.5} />
      <path d="M2.5 8h11" />
      <path d="M8 2.5c1.8 2 1.8 9 0 11" />
      <path d="M8 2.5c-1.8 2-1.8 9 0 11" />
    </Icon>
  );
}

export function UserIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <circle cx={8} cy={6} r={2.25} />
      <path d="M4 14.5c.9-2.3 2.6-3.5 4-3.5s3.1 1.2 4 3.5" />
    </Icon>
  );
}

export function UsersIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <circle cx={6} cy={6} r={2} />
      <circle cx={11} cy={6.75} r={1.75} />
      <path d="M2.75 14.5c.8-2 2.3-3 3.25-3s2.45 1 3.25 3" />
      <path d="M9.5 14.5c.5-1.3 1.6-2.25 2.75-2.25 1.15 0 2.25.95 2.75 2.25" />
    </Icon>
  );
}

export function XIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M5 5l6 6" />
      <path d="M11 5l-6 6" />
    </Icon>
  );
}

export function CheckIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M4 8l2.5 2.5L12 5" />
    </Icon>
  );
}

export function WarningIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M8 2.5l6 11H2L8 2.5z" />
      <path d="M8 6.5v3" />
      <path d="M8 11.75h.01" />
    </Icon>
  );
}

export function HelpIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M6.5 6.25a1.5 1.5 0 1 1 2.5 1.13c-.6.45-1 .8-1 1.62" />
      <path d="M8 12h.01" />
      <circle cx={8} cy={8} r={5.5} />
    </Icon>
  );
}

export function TrashIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M5 4.5h6" />
      <path d="M6 4.5V3.5A1 1 0 0 1 7 2.5h2a1 1 0 0 1 1 1v1" />
      <path d="M4.5 4.5l.5 10A1 1 0 0 0 6 15.5h4a1 1 0 0 0 1-1l.5-10" />
      <path d="M7 7v6" />
      <path d="M9 7v6" />
    </Icon>
  );
}

export function EyeIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M2.5 8s2-3.5 5.5-3.5S13.5 8 13.5 8s-2 3.5-5.5 3.5S2.5 8 2.5 8z" />
      <circle cx={8} cy={8} r={1.5} />
    </Icon>
  );
}

export function EyeOffIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M3 3l10 10" />
      <path d="M2.5 8s2-3.5 5.5-3.5c1.2 0 2.2.4 3 1" />
      <path d="M13.5 8s-2 3.5-5.5 3.5c-1.2 0-2.2-.4-3-1" />
      <path d="M7.25 7.25a1.5 1.5 0 0 0 1.5 1.5" />
    </Icon>
  );
}

export function CommentIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M3 4.5h10v6.5A2 2 0 0 1 11 13.5H7l-2.5 2v-2H5A2 2 0 0 1 3 11V4.5z" />
    </Icon>
  );
}

export function ChartIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M3 13.5h11" />
      <path d="M4.5 13.5v-4" />
      <path d="M7.5 13.5v-6" />
      <path d="M10.5 13.5v-2" />
      <path d="M4.5 9l3-2 3 3 2-1" />
    </Icon>
  );
}

export function ImageIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <rect x={3} y={4} width={10} height={9} rx={1.5} />
      <path d="M5 11l2.25-2.25L10.5 12" />
      <circle cx={6} cy={6.5} r={0.75} />
    </Icon>
  );
}

export function PuzzleIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M6 6.25a1.25 1.25 0 1 1 2.5 0V7h2.5v2.5H10.25a1.25 1.25 0 1 1 0 2.5H11v2.5H8.5v-.75a1.25 1.25 0 1 0-2.5 0V15H3.5v-2.5H4.25a1.25 1.25 0 1 0 0-2.5H3.5V7H6v-.75z" />
    </Icon>
  );
}

export function WindowIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <rect x={3} y={4} width={10} height={9} rx={1.5} />
      <path d="M3 6.5h10" />
      <path d="M5 5.25h.01" />
      <path d="M7 5.25h.01" />
    </Icon>
  );
}

export function ArrowUpIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M8 13V3" />
      <path d="M5 6l3-3 3 3" />
    </Icon>
  );
}

export function ArrowDownIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M8 3v10" />
      <path d="M5 10l3 3 3-3" />
    </Icon>
  );
}

export function ArrowLeftIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M13 8H3" />
      <path d="M6 5l-3 3 3 3" />
    </Icon>
  );
}

export function ArrowRightIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M3 8h10" />
      <path d="M10 5l3 3-3 3" />
    </Icon>
  );
}

export function ArrowLeftRightIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M3.5 8h9" />
      <path d="M5.5 6l-2 2 2 2" />
      <path d="M10.5 6l2 2-2 2" />
    </Icon>
  );
}

export function ArrowUpDownIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M8 3v10" />
      <path d="M6 5l2-2 2 2" />
      <path d="M6 11l2 2 2-2" />
    </Icon>
  );
}

export function ReturnIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M12.5 5.5H6A2.5 2.5 0 0 0 3.5 8v2.5" />
      <path d="M5.5 8.5l-2 2 2 2" />
    </Icon>
  );
}

export function PlusIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M8 4v8" />
      <path d="M4 8h8" />
    </Icon>
  );
}

export function MinusIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M4 8h8" />
    </Icon>
  );
}

export function DivideIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M4 8h8" />
      <path d="M8 5.25h.01" />
      <path d="M8 10.75h.01" />
    </Icon>
  );
}

export function ShuffleIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M3 5h2.5l2 2L11 7.5h2" />
      <path d="M13 5.5l2 2-2 2" />
      <path d="M3 11h2.5l2-2" />
      <path d="M11 10.5h2" />
      <path d="M13 8.5l2 2-2 2" />
    </Icon>
  );
}

export function PlayIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <path d="M6 4.5v7l6-3.5-6-3.5z" />
    </Icon>
  );
}

export function StopIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <rect x={5} y={5} width={6} height={6} rx={1} />
    </Icon>
  );
}

export function RecordIcon(props: RibbonSvgIconProps) {
  return (
    <Icon {...props}>
      <circle cx={8} cy={8} r={3} />
    </Icon>
  );
}
