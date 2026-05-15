export interface IHashtag {
  Id: number;
  Title: string;
  Description?: string;
  Category?: string;
}

export interface IHashtagGroup {
  name: string;
  items: IHashtag[];
}

const UNCATEGORIZED = 'Khác';

export function groupHashtagsByCategory(tags: IHashtag[]): IHashtagGroup[] {
  const buckets = new Map<string, IHashtag[]>();
  for (const t of tags) {
    const key = (t.Category || '').trim() || UNCATEGORIZED;
    const existing = buckets.get(key);
    if (existing) existing.push(t); else buckets.set(key, [t]);
  }
  const names = Array.from(buckets.keys()).sort((a, b) => {
    if (a === UNCATEGORIZED && b !== UNCATEGORIZED) return 1;
    if (b === UNCATEGORIZED && a !== UNCATEGORIZED) return -1;
    return a.localeCompare(b);
  });
  return names.map(name => ({
    name,
    items: (buckets.get(name) || []).slice().sort((x, y) => x.Title.localeCompare(y.Title))
  }));
}

export interface IDocument {
  Id: number;
  Name: string;
  ServerRelativeUrl: string;
  Created: string;
  Modified: string;
  CreatedBy: string;
  SizeKB: number;
  Hashtags: IHashtag[];
}

export interface IUploadResult {
  success: boolean;
  fileName: string;
  error?: string;
}

export type SearchMode = 'any' | 'all';
