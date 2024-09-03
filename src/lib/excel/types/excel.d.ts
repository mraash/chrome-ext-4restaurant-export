import { CellKey } from '../CellKey';
import { CellValue } from '../CellValue';

// CellKey
export type CellColumnType = string | number;
export type CellRowType = number;

/**
 * "B1", ["B", 1], [2, 1]
 */
export type CellKeyPrimitivesType = string | [string, number] | [number, number];

export type CellKeyType = CellKey | CellKeyPrimitivesType;

// CellValue
export type CellValueType = CellValue | string | number;
