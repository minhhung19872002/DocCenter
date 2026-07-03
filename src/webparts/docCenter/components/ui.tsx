import * as React from 'react';
import { IHashtag } from '../services/types';
import styles from './DocCenter.module.scss';

// Bảng màu theo nhóm hashtag (theo design "Trung tâm Tài liệu")
export interface ICatPalette {
  solid: string;
  text: string;
  bg: string;
  bd: string;
}

const PALETTES: { [key: string]: ICatPalette } = {
  congty: { solid: '#1866b8', text: '#1866b8', bg: '#eef5fd', bd: '#cfe3f7' },
  loai:   { solid: '#c05613', text: '#9a5b10', bg: '#fdf3e3', bd: '#f3dfba' },
  phong:  { solid: '#00913f', text: '#00913f', bg: '#e4f4e9', bd: '#bfe4cc' },
  khac:   { solid: '#7c3aed', text: '#5f2f96', bg: '#f3ecfa', bd: '#e2d3f2' },
  teal:   { solid: '#0f766e', text: '#0f766e', bg: '#e6f7f5', bd: '#bfe8e3' },
  rose:   { solid: '#be185d', text: '#9d1a52', bg: '#fdeef4', bd: '#f6cede' }
};

// Nhóm chưa biết tên sẽ được gán màu ổn định từ danh sách này
const FALLBACK_KEYS = ['congty', 'loai', 'phong', 'teal', 'rose'];

function normalize(s: string): string {
  return s
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/đ/gi, 'd')
    .trim()
    .toLowerCase();
}

const KNOWN_NAMES: { [key: string]: string } = {
  'cong ty': 'congty',
  'loai tai lieu': 'loai',
  'loai': 'loai',
  'phong ban': 'phong',
  'phong': 'phong',
  'khac': 'khac',
  '': 'khac'
};

export function getCategoryPalette(category?: string): ICatPalette {
  const key = normalize(category || '');
  const known = KNOWN_NAMES[key];
  if (known) return PALETTES[known];
  let hash = 0;
  for (let i = 0; i < key.length; i++) hash = ((hash << 5) - hash + key.charCodeAt(i)) | 0;
  return PALETTES[FALLBACK_KEYS[Math.abs(hash) % FALLBACK_KEYS.length]];
}

export function chipStyle(p: ICatPalette, selected: boolean, small?: boolean): React.CSSProperties {
  const base: React.CSSProperties = {
    display: 'inline-flex',
    alignItems: 'center',
    fontSize: small ? 11 : 12.5,
    fontWeight: selected ? 700 : 600,
    padding: small ? '3px 10px' : '5px 14px',
    borderRadius: 18,
    cursor: 'pointer',
    userSelect: 'none',
    whiteSpace: 'nowrap',
    transition: 'all .12s'
  };
  if (selected) {
    return { ...base, background: p.solid, color: '#fff', border: `1px solid ${p.solid}` };
  }
  return { ...base, background: p.bg, color: p.text, border: `1px solid ${p.bd}` };
}

interface ICatChipProps {
  tag: IHashtag;
  selected?: boolean;
  small?: boolean;
  readonly?: boolean;
  showRemove?: boolean;
  prefix?: string;
  onClick?: () => void;
}

export const CatChip: React.FunctionComponent<ICatChipProps> = props => {
  const p = getCategoryPalette(props.tag.Category);
  const style = chipStyle(p, !!props.selected, props.small);
  if (props.readonly) style.cursor = 'default';
  return (
    <span
      style={style}
      title={props.tag.Description || undefined}
      onClick={props.readonly ? undefined : props.onClick}
    >
      {props.prefix || ''}#{props.tag.Title}{props.showRemove ? ' ✕' : ''}
    </span>
  );
};

export const GroupLabel: React.FunctionComponent<{ name: string }> = props => {
  const p = getCategoryPalette(props.name);
  return (
    <span className={styles.groupLabel} style={{ color: p.text }}>
      {props.name}
    </span>
  );
};

// Ô màu hiển thị đuôi file (thay cho icon Office)
const EXT_COLORS: { [ext: string]: [string, string] } = {
  pdf:  ['#fbe9e7', '#c8342b'],
  doc:  ['#e7effa', '#1866b8'],
  docx: ['#e7effa', '#1866b8'],
  xls:  ['#e7f5ec', '#0f7b3e'],
  xlsx: ['#e7f5ec', '#0f7b3e'],
  csv:  ['#e7f5ec', '#0f7b3e'],
  ppt:  ['#fdf3e3', '#c05613'],
  pptx: ['#fdf3e3', '#c05613']
};
const EXT_DEFAULT: [string, string] = ['#eef1ea', '#68766c'];

export function extOf(fileName: string): string {
  const dot = fileName.lastIndexOf('.');
  return dot >= 0 ? fileName.substring(dot + 1).toLowerCase() : '';
}

export const ExtBadge: React.FunctionComponent<{ fileName: string; size?: number }> = props => {
  const ext = extOf(props.fileName);
  const label = (ext || 'FILE').toUpperCase().substring(0, 4);
  const c = EXT_COLORS[ext] || EXT_DEFAULT;
  const size = props.size || 38;
  return (
    <span
      className={styles.extBadge}
      style={{
        width: size,
        height: size,
        background: c[0],
        color: c[1],
        fontSize: label.length > 3 ? 8 : 9.5
      }}
    >
      {label}
    </span>
  );
};

// Segmented control (chuyển chế độ VÀ/HOẶC, sub-tab Quản trị)
export interface ISegOption {
  key: string;
  label: string;
}

interface ISegmentedProps {
  options: ISegOption[];
  value: string;
  variant?: 'admin';
  onChange: (key: string) => void;
}

export const Segmented: React.FunctionComponent<ISegmentedProps> = props => (
  <div className={props.variant === 'admin' ? styles.segWrapAdmin : styles.segWrap}>
    {props.options.map(o => {
      const active = o.key === props.value;
      const cls = props.variant === 'admin'
        ? `${styles.segBtn} ${active ? styles.segBtnActiveAdmin : ''}`
        : `${styles.segBtn} ${active ? styles.segBtnActive : ''}`;
      return (
        <div key={o.key} className={cls} onClick={() => props.onChange(o.key)}>
          {o.label}
        </div>
      );
    })}
  </div>
);
