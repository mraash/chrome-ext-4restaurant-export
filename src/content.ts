import * as XLSX from 'xlsx';

chrome.runtime.onMessage.addListener((message) => {
    if (message.action !== 'ACTION_EXPORT') {
        return;
    }

    const sheet = extractTextWithTableReplacement(document.body);
    const workbook = convertSheetDataToWorkbook(sheet);
    saveSheetFile(workbook);
});

function extractTextWithTableReplacement(element: Element, result: any[] = []) {
    if (element.nodeType === Node.TEXT_NODE) {
        let trimmedText = element.textContent?.trim();

        if (trimmedText && trimmedText.length > 0) {
            // if (trimmedText.match(/^Valdes priekšsēdētājs/)) {
            //     trimmedText = trimmedText.replace('Valdes priekšsēdētājs', 'Noliktavas parzinis');
            // }

            if (trimmedText.match(/^Nodokļu maksātāja/)) {
                result.push([null]);
            }
            if (trimmedText.match(/^Produktu pieprasījums/)) {
                result.push([null]);
            }
            if (trimmedText.match(/^Pakalpojuma/)) {
                result.push([null]);
            }
            if (trimmedText.match(/^Apstiprināja/)) {
                result.push([null]);
            }
            if (trimmedText.match(/^Virtuves vadītāja/)) {
                result.push([null]);
            }
            if (trimmedText.match(/^Valdes priekšsēdētājs/)) {
                result.push([null]);
            }

            result.push([trimmedText]);
        }
    }
    else if (element.nodeType === Node.ELEMENT_NODE) {
        if (element.nodeName === 'TABLE') {
            if (element.querySelectorAll('tr').length > 1) {
                result.push([null]);

                const firstRow = result.length;
                const tableSheet = convertTableObjectToSheetData(element as HTMLTableElement, firstRow);

                tableSheet.forEach((item) => {
                    result.push(item);
                });
            }
        }
        else {
            element.childNodes.forEach(child => extractTextWithTableReplacement(child as Element, result));
        }
    }

    return result;
}

function convertSheetDataToWorkbook(sheetData: Array<any>): XLSX.WorkBook {
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Report');

    return workbook;
}

function saveSheetFile(workbook: XLSX.WorkBook): void {
    XLSX.writeFile(workbook, 'Report.xlsx');
}

function convertTableObjectToSheetData(table: HTMLTableElement, startRow: number) {
    const tableData = convertHtmlTableToObject(table);

    const sheetData: Array<Array<XLSX.CellObject | string | number>> = [
        [
            'Kods',
            'Nosaukums',
            'Mērvienība',
            'Piezīmes',
            '1. Brokastis',
            '2. Pusdienas',
            '3. Launags',
            '4. Vakariņas',
            'Kopā',
        ],
    ];

    for (let rowIndex = 0; rowIndex < tableData.length; rowIndex++) {
        const currentRowIndex = rowIndex + 2 + startRow;

        const code = tableData[rowIndex].code;
        const name = tableData[rowIndex].name;
        const unit = tableData[rowIndex].unit;
        const notes = tableData[rowIndex].notes;
        const breakfest = Number(tableData[rowIndex].breakfest);
        const lunck = Number(tableData[rowIndex].lunck);
        const secondsLunch = Number(tableData[rowIndex].secondsLunch);
        const dinner = Number(tableData[rowIndex].dinner);
        const total = Number(tableData[rowIndex].total);

        const { breakfestFormula, luncFormula, secondsLunchFormula, dinnerFormula } = createQuantityColumnFormulas(
            currentRowIndex,
            total,
            breakfest,
            lunck,
            secondsLunch,
            dinner,
        );

        sheetData.push([
            { t: 's', v: code, },
            name,
            unit,
            notes,
            { t: 'n', v: breakfest, f: breakfestFormula },
            { t: 'n', v: lunck, f: luncFormula },
            { t: 'n', v: secondsLunch, f: secondsLunchFormula },
            { t: 'n', v: dinner, f: dinnerFormula },
            { t: 'n', v: total, },
        ]);
    }

    return sheetData;
};

function createQuantityColumnFormulas(
    row: number,
    totalQuantity: number,
    breakfestQuantity: number,
    lunckQuantity: number,
    secondsLunchQuantity: number,
    dinnerQuantity: number,
) {
    const breakfestProcent = breakfestQuantity / totalQuantity;
    const lunckProcent = lunckQuantity / totalQuantity;
    const secondsLunchProcent = secondsLunchQuantity / totalQuantity;
    const dinnerProcent = dinnerQuantity / totalQuantity;

    let breakfestFormula = `ROUND(I${row} * ${breakfestProcent}, 3)`;
    let luncFormula = `ROUND(I${row} * ${lunckProcent}, 3)`;
    let secondsLunchFormula = `ROUND(I${row} * ${secondsLunchProcent}, 3)`;
    let dinnerFormula = `ROUND(I${row} * ${dinnerProcent}, 3)`;

    let lastNotNull = '';

    if (breakfestQuantity > 0)  lastNotNull = 'b';
    if (lunckQuantity > 0)  lastNotNull = 'l';
    if (secondsLunchQuantity > 0)  lastNotNull = 's';
    if (dinnerQuantity > 0)  lastNotNull = 'd';

    if (lastNotNull === 'b') breakfestFormula = `I${row} - F${row} - G${row} - H${row}`;
    if (lastNotNull === 'l') luncFormula = `I${row} - E${row} - G${row} - H${row}`;
    if (lastNotNull === 's') secondsLunchFormula = `I${row} - E${row} - F${row} - H${row}`;
    if (lastNotNull === 'd') dinnerFormula = `I${row} - E${row} - F${row} - G${row}`;

    return {
        breakfestFormula,
        luncFormula,
        secondsLunchFormula,
        dinnerFormula,
    };
}

function convertHtmlTableToObject($table: HTMLElement) {
    const table = [];
    const $rows = $table.querySelectorAll('tr');

    for (let i = 0; i < $rows.length; i++) {
        const $cells = $rows[i].querySelectorAll('td');

        if ($cells.length !== 9) {
            continue;
        }

        table.push({
            code: $cells[0].textContent!,
            name: $cells[1].textContent!,
            unit: $cells[2].textContent!,
            notes: $cells[3].textContent!,
            breakfest: $cells[4].textContent!.replace(',', '.'),
            lunck: $cells[5].textContent!.replace(',', '.'),
            secondsLunch: $cells[6].textContent!.replace(',', '.'),
            dinner: $cells[7].textContent!.replace(',', '.'),
            total: $cells[8].textContent!.replace(',', '.'),
        });
    }

    return table;
}
