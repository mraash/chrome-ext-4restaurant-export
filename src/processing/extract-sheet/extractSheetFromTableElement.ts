import XLSX from 'xlsx-js-style';
import { Sheet, SheetCell, SheetRow } from '../../types';
import { cell } from './processCell';
import { userSettings } from '../../user-settings';
import { setPercentFormulas } from './formula/setProcentFormulas';

export const extractSheetFromTableElement = (
    table: HTMLTableElement,
    startRow: number,
): Sheet => {
    const $rows = table.querySelectorAll('tr');
    const resultRows: SheetRow[] = [];
    const tableHeader: SheetRow = [];

    // Todo: Handle cases where there are other elements besides tr th and td.

    $rows.forEach(($row, index) => {
        const currentRow = startRow + resultRows.length;

        if (index === 0) {
            const $headers = $row.querySelectorAll('th');

            if ($headers.length === 0) {
                // Todo: Handle exception
            }

            $headers.forEach($header => {
                tableHeader.push(processTextElement($header));                
            });

            resultRows.push(tableHeader);

            return;
        }

        const $cells = $row.querySelectorAll('td');
        const row: SheetRow = [];

        $cells.forEach($cell => {
            row.push(processTextElement($cell));
        });

        for (let i = 0; i < userSettings.applyPercentFormula.length; i++) {
            const formulaSettings = userSettings.applyPercentFormula[i];

            if (formulaSettings.tableWithHeaders.length !== tableHeader.length) {
                continue;
            }

            for (let i = 0; i < formulaSettings.tableWithHeaders.length; i++) {
                if (tableHeader.find(headerCell => formulaSettings.tableWithHeaders.includes(headerCell.v as string)) === undefined) {
                    continue;
                }
            }

            const { totalCellHeader, percentCellHeaders } = formulaSettings.options;

            const totalCellIndex = tableHeader.findIndex(headerCell => headerCell.v == totalCellHeader);

            const percentCellIndexes: number[] = percentCellHeaders.map(
                headerValue => tableHeader.findIndex(header => header.v == headerValue)
            );

            setPercentFormulas(row, {
                totalCellIndex: totalCellIndex,
                percentCellIndexes: percentCellIndexes,
                rowNumber: currentRow
            });

            resultRows.push(row);

            break;
        }
    });

    return resultRows;
};

function processTextElement(element: Element): SheetCell {
    const trimmedText = element.textContent?.trim();
    const isElementEmpty = trimmedText === undefined || trimmedText === '';

    let result = isElementEmpty ? cell.getCell('') : cell.getCell(cell.getValue(trimmedText));

    // TODO: Styles can be moved to a separate section of code
    const borderStyle = {
        style: 'thin',
        color: {
            rgb: '000000',
        },
    };

    result.s = {
        border: {
            top: borderStyle,
            right: borderStyle,
            bottom: borderStyle,
            left: borderStyle,
        },
    };

    return result;
}
