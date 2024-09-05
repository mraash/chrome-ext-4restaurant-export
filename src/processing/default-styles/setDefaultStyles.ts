import { Sheet } from '../../types';
import { userSettings } from '../../user-settings';

export const setDefaultStyles = (sheetData: Sheet): Sheet => {
    for (let rowIndex = 0; rowIndex < sheetData.length; rowIndex++) {
        const row = sheetData[rowIndex];

        for (let colIndex = 0; colIndex < row.length; colIndex++) {
            const cell = row[colIndex];
            
            cell.s = cell?.s ?? {};
            cell.s.font = cell.s?.font ?? {};

            cell.s.font.name = cell.s.font?.name ?? userSettings.font;
        }
    }

    return sheetData;
};
