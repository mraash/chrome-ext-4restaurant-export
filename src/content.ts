import * as XLSX from 'xlsx-js-style';

chrome.runtime.onMessage.addListener((message) => {
    if (message.action !== 'ACTION_EXPORT') {
        return;
    }

    const sheet = extractTextWithTableReplacement(document.body);
    const styledSheet = setDefaultStylesToSheetData(sheet);
    const workbook = convertSheetDataToWorkbook(styledSheet);
    saveSheetFile(workbook);
});

type SheetData = Array<Array<XLSX.CellObject>>;

function extractTextWithTableReplacement(element: Element, result: any[] = []): SheetData {
    if (element.nodeType === Node.TEXT_NODE) {
        let trimmedText = element.textContent?.trim();

        if (trimmedText && trimmedText.length > 0) {
            if (trimmedText.match(/^Valdes priekšsēdētājs/)) {
                trimmedText = trimmedText.replace('Valdes priekšsēdētājs', 'Noliktavas pārzinis');
            }

            if (trimmedText.match(/^Nodokļu maksātāja/)) {
                result.push([null]);
            }
            if (trimmedText.match(/^Produktu pieprasījums/)) {
                result.push([null]);
            }
            if (trimmedText.match(/^Pakalpojuma/)) {
                result.push([null]);
            }
            if (trimmedText.match(/^Virtuves vadītāja/)) {
                result.push([null]);
            }
            if (trimmedText.match(/^Noliktavas pārzinis/)) {
                result.push([null]);
            }

            const cell: XLSX.CellObject = { t: 's', v: trimmedText };

            if (trimmedText.match(/^Produktu pieprasījums/)) {
                cell.s = {
                    font: {
                        bold: true,
                    }
                }
            }

            result.push([ cell ]);
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

function setDefaultStylesToSheetData(sheetData: SheetData): SheetData {
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
            cell.s.border = cell.s?.border ?? {};
            cell.s.font = cell.s?.font ?? {};

            cell.s.border.top = cell.s.border?.top ?? borderStyle;
            cell.s.border.right = cell.s.border?.right ?? borderStyle;
            cell.s.border.bottom = cell.s.border?.bottom ?? borderStyle;
            cell.s.border.left = cell.s.border?.left ?? borderStyle;

            cell.s.font.name = cell.s.font?.name ?? 'Times New Roman';
        }
    }

    return sheetData;
}

function convertSheetDataToWorkbook(sheetData: SheetData): XLSX.WorkBook {
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Report');

    return workbook;
}

function saveSheetFile(workbook: XLSX.WorkBook): void {
    XLSX.writeFile(workbook, 'Report.xlsx', { sheetStubs: true });
}

function convertTableObjectToSheetData(table: HTMLTableElement, startRow: number) {
    const tableData = convertHtmlTableToObject(table);

    const sheetData: SheetData = [];

    const header = [
        'Kods',
        'Nosaukums',
        'Mērvienība',
        'Piezīmes',
        '1. Brokastis',
        '2. Pusdienas',
        '3. Launags',
        '4. Vakariņas',
        'Kopā',
    ].map((value) => ({ t: 's' as XLSX.ExcelDataType, v: value, s: { font: { bold: true, textAlign: 'center'} } }));

    sheetData.push(header);

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
            { t: 's', v: code },
            { t: 's', v: name },
            { t: 's', v: unit },
            { t: 's', v: notes },
            { t: 'n', v: breakfest, f: breakfestFormula },
            { t: 'n', v: lunck, f: luncFormula },
            { t: 'n', v: secondsLunch, f: secondsLunchFormula },
            { t: 'n', v: dinner, f: dinnerFormula },
            { t: 'n', v: total },
        ]);
    }

    const borderStyle = {
        style: 'thin',
        color: {
            rgb: '000000',
        },
    };

    const s = {
        border: {
            top: borderStyle,
            right: borderStyle,
            bottom: borderStyle,
            left: borderStyle,
        }
    };

    sheetData.forEach((row) => {
        row.forEach((cell) => {
            if (!cell.s) {
                cell.s = {};
            }

            cell.s.border = s.border;
        });

        row.push({
            t: 's',
            v: '',
            s: {
                border: {
                    left: borderStyle,
                }
            }
        });
    });

    sheetData.push([
        { t: 's', v: '', s: { border: { top: borderStyle } } },
        { t: 's', v: '', s: { border: { top: borderStyle } } },
        { t: 's', v: '', s: { border: { top: borderStyle } } },
        { t: 's', v: '', s: { border: { top: borderStyle } } },
        { t: 's', v: '', s: { border: { top: borderStyle } } },
        { t: 's', v: '', s: { border: { top: borderStyle } } },
        { t: 's', v: '', s: { border: { top: borderStyle } } },
        { t: 's', v: '', s: { border: { top: borderStyle } } },
        { t: 's', v: '', s: { border: { top: borderStyle } } },
    ]);

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
