export interface ColumnType {
  name: string;
  type: 'text' | 'numeric' | 'date';
}

export type DataRow = Record<string, any>;

export interface ColumnMapping {
  dateCol: string;
  categoryCol: string;
  revenueCol: string;
  costCol: string;
  currency?: string;
}

export interface KPIData {
  totalRevenue: number;
  totalCost: number;
  totalProfit: number;
  roi: number;
  totalRows: number;
  uniqueCategories: number;
  hasCost: boolean;
}

export interface FilterState {
  dateStart: string;
  dateEnd: string;
  categories: string[];
}

export type SortDirection = 'asc' | 'desc' | null;

export interface SortState {
  column: string | null;
  direction: SortDirection;
}

export type ScreenState = 'upload' | 'mapping' | 'dashboard';
export type ThemeMode = 'dark' | 'light';
