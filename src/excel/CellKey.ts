import { CellKeyPrimitivesType, CellKeyType } from './types/excel';

export class CellKey {
    private row: number;
    private column: number;

    public constructor(column: number, row: number) {
        this.row = row;
        this.column = column;
    }

    public static fromPrimitives(key: CellKeyPrimitivesType): CellKey {
        if (key instanceof CellKey) {
            return key;
        }

        let column: string | number;
        let row: number;

        if (typeof key === 'string') {
            const keyObject = this.splitStringKey(key);

            column = this.columnToNumber(keyObject.column),
            row = keyObject.row
        }
        else {
            column = typeof key[0] === 'string' ? this.columnToNumber(key[0]) : key[0];
            row = key[1];
        }

        return new this(column, row);
    }

    public toString(): string {
        return CellKey.numberToColumn(this.column) + this.row;
    }

    public getColumn(): number {
        return this.column;
    }

    public getRow(): number {
        return this.row;
    }

    public static numberToColumn(n: number): string {
        if (n <= 0) {
            return '';
        }

        return this.numberToColumn(Math.floor(n / 26)) + String.fromCharCode(65 + (n % 26));
    }

    public static columnToNumber(column: string): number {
        if (column === '') {
            return 0;
        }

        return (column.charCodeAt(0) - 64) * Math.pow(26, column.length - 1) + this.columnToNumber(column.slice(1));
    }

    public static splitStringKey(key: string): { column: string; row: number } {
        if (!this.isStringKeyValid(key)) {
            throw new Error("Invalid cell key format");
        }

        let splitIndex = 0;

        while (splitIndex < key.length && isNaN(Number(key[splitIndex]))){
            splitIndex++;
        }

        const column = key.slice(0, splitIndex);
        const row = Number(key.slice(splitIndex));

        return { column, row };
    }

    public static isStringKeyValid(key: string): boolean {
        const regex = /^[A-Z]+[0-9]+$/i;

        return regex.test(key);
    }
}
