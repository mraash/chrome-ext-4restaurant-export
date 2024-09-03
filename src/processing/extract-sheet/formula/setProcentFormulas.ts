import { CellKey } from '../../../lib/excel/CellKey';
import { SheetRow } from '../../../types';

export const setPercentFormulas = (row: SheetRow, options: Options) => {
    const totalCellKey = getColumnByRowIndex(options.totalCellIndex) + options.rowNumber;
    const totalCellValue = row[options.totalCellIndex].v as number;

    options.percentCellIndexes.forEach(rowIndex => {
        const percentCellValue = row[rowIndex].v as number;
        const percent = percentCellValue / totalCellValue;

        row[rowIndex].f = `ROUND(${totalCellKey} * ${percent}, 3)`;
    });

    for (let i = options.percentCellIndexes.length - 1; i >= 0; i--) {
        const rowIndex = options.percentCellIndexes[i];
        const percentCellValue = row[rowIndex].v as number;

        if (percentCellValue > 0) {
            let formula = `${totalCellKey}`;

            options.percentCellIndexes.forEach(otherRowIndex => {
                if (otherRowIndex === rowIndex) {
                    return;
                }

                formula += ` - ${getColumnByRowIndex(otherRowIndex)}${options.rowNumber}`;
            });

            row[rowIndex].f = formula;

            break;
        }
    }
};

export type Options = {
    percentCellIndexes: number[],
    totalCellIndex: number
    rowNumber: number,
};

function getColumnByRowIndex(index: number): string {
    return CellKey.numberToColumn(index + 1);
}
