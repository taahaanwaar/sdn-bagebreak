
export type ExcelRow = (string | number | boolean | null | undefined)[];

export interface ColumnMeta {
  wpx: number; // Width in pixels
}

export interface RowMeta {
  hpx: number; // Height in pixels
}

export interface PageGroup {
  id: string;
  rows: ExcelRow[];
  rowMetas: (RowMeta | undefined)[];
  siteHeader: {
    sdnId: string;
    sdnType: string;
    sdnDate: string;
    reqDate: string;
    szsPob: string;
    clientPob: string;
    location: string;
    siteName: string;
  };
}
