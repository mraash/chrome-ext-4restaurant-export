import XLSX from 'xlsx-js-style';

export type Sheet = SheetRow[];
export type SheetRow = SheetCell[];
export type SheetCell = XLSX.CellObject;
