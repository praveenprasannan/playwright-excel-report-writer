export type ColumnConfig = {
  key: string;
  header: string;
  width?: number;
  type?: 'text' | 'image';
};

export type SheetConfig = {
  strategy: 'single' | 'by-browser';
  name?: string;
  namePattern?: string; // e.g. "Results - {browser}"
};

export type FileConfig = {
  directory: string;
  fileName: string;
};

export type ImageConfig = {
  columnKey: string;
  showOnlyOnFail?: boolean;
  width: number;
  height: number;
};

export type StyleConfig = {
  headerBold?: boolean;
  headerSize?: number;
  passFill?: string;
  failFill?: string;
  borderStyle?: 'thin' | 'medium';
};

export type ExcelWriterConfig = {
  file: FileConfig;
  sheet: SheetConfig;
  columns: ColumnConfig[];
  rowKey: string;
  image?: ImageConfig;
  styles?: StyleConfig;
  lockTimeoutMs?: number;
};
