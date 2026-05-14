export interface IHashtag {
  Id: number;
  Title: string;
}

export interface IDocument {
  Id: number;
  Name: string;
  ServerRelativeUrl: string;
  Modified: string;
  ModifiedBy: string;
  SizeKB: number;
  Hashtags: IHashtag[];
}

export interface IUploadResult {
  success: boolean;
  fileName: string;
  error?: string;
}

export type SearchMode = 'any' | 'all';
