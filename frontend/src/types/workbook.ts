export type CellKind = 'empty' | 'formula' | 'string' | 'number' | 'bool' | 'date' | 'error' | string;

export type DataValidation = {
  type: string | null;
  operator: string | null;
  formula1: string | null;
  formula2: string | null;
  allow_blank: boolean | null;
  show_dropdown: boolean | null;
  sqref: string | null;
  list_values?: unknown[];
};

export type WorkbookCell = {
  address: string;
  kind: CellKind;
  value?: unknown;
  formula?: string;
  /** Last value Excel stored for this formula (data_only read); not recomputed in the browser. */
  cached_value?: unknown;
  number_format?: string;
  validation?: DataValidation;
};

export type WorkbookSheet = {
  name: string;
  sheet_state?: string;
  max_row: number;
  max_column: number;
  merged_ranges: string[];
  data_validations?: DataValidation[];
  cells: WorkbookCell[];
};

export type WorkbookParseMeta = {
  region_max_column_letter?: string;
  region_max_row?: number;
  [key: string]: unknown;
};

export type WorkbookDump = {
  workbook: string;
  data_only: boolean;
  parse?: WorkbookParseMeta;
  sheet_count: number;
  sheets: WorkbookSheet[];
};
