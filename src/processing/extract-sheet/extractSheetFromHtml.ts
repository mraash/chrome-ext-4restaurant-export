import { Sheet, SheetCell, SheetRow } from '../../types';
import { userSettings } from '../../user-settings';
import { extractSheetFromTableElement } from './extractSheetFromTableElement';
import { cell } from './processCell';

// TODO: The variable previousElement is needed only to avoid adding double
//   line breaks in excel if there are several br elements in a row in HTML.
//   The solution with previousElement is terrible and should be fixed.
let previousElement: Element | null;

export const extractSheetFromHtml = (element: Element, rows: SheetRow[] = []): Sheet => {
    const isFirstIteration = rows.length === 0;

    if (isFirstIteration) {
        previousElement = null;
    }

    if (element.nodeType === Node.TEXT_NODE) {
        const text = processTextElement(element);

        if (text !== null) {
            previousElement = element;
            rows.push([text]);
        }

        return rows;
    }

    if (element.nodeName === 'BR') {
        if (previousElement?.nodeName !== 'BR') {
            rows.push([]);
        }

        previousElement = element;
        return rows;
    }

    if (element.nodeName === 'TABLE') {
        // TODO: set isTableEmpty correctly
        const isTableEmpty = element.querySelectorAll('tr').length <= 1;

        if (isTableEmpty && userSettings.ignoreEmptyTable) {
            return rows;
        }

        const tableSheet = extractSheetFromTableElement(element as HTMLTableElement, rows.length + 1);

        tableSheet.forEach((row) => {
            rows.push(row);
        });

        previousElement = element;
        return rows;
    }

    element.childNodes.forEach(child => extractSheetFromHtml(child as Element, rows));

    previousElement = element;
    return rows;
};

function processTextElement(element: Element): SheetCell | null {
    let trimmedText = element.textContent?.trim();

    if (trimmedText === undefined || trimmedText === '') {
        return null;
    }

    return cell.getCell(cell.getValue(trimmedText));
}
