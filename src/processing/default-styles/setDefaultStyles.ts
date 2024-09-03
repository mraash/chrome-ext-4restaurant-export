import { Sheet } from '../../types';

export const setDefaultStyles = (sheetData: Sheet): Sheet => {
    const MIN_ROWS = 100;
    const MIN_COLUMNS = 50;

    const borderStyle = {
        style: 'thin',
        color: {
            rgb: 'FFFFFF',
        },
    }

    for (let rowIndex = 0; rowIndex < sheetData.length || rowIndex < MIN_ROWS; rowIndex++) {
        const isRowEmpty = !sheetData[rowIndex];

        if (isRowEmpty) {
            sheetData.push([]);
        }

        const row = sheetData[rowIndex];

        for (let colIndex = 0; colIndex < row.length || colIndex < MIN_COLUMNS; colIndex++) {
            const isCellEmpty = !row[colIndex];

            if (isCellEmpty) {
                sheetData[rowIndex][colIndex] = { t: 's', v: '' };
            }

            const cell = row[colIndex];
            
            cell.s = cell?.s ?? {};
            // cell.s.border = cell.s?.border ?? {};
            cell.s.font = cell.s?.font ?? {};

            // cell.s.border.top = cell.s.border?.top ?? borderStyle;
            // cell.s.border.right = cell.s.border?.right ?? borderStyle;
            // cell.s.border.bottom = cell.s.border?.bottom ?? borderStyle;
            // cell.s.border.left = cell.s.border?.left ?? borderStyle;

            cell.s.font.name = cell.s.font?.name ?? 'Times New Roman';
        }
    }

    return sheetData;
};
