import * as XLSX from 'xlsx-js-style';
import { extractSheetFromHtml } from './processing/extract-sheet/extractSheetFromHtml';
import { Sheet } from './types';
import { setDefaultStyles } from './processing/default-styles/setDefaultStyles';
import { userSettings } from './user-settings';

chrome.runtime.onMessage.addListener((message) => {
    if (message.action !== 'ACTION_EXPORT') {
        return;
    }

    const sheet = extractSheetFromHtml(document.body);
    const styledSheet = setDefaultStyles(sheet);

    const workbook = convertSheetDataToWorkbook(styledSheet);
    saveSheetFile(workbook);
});

function convertSheetDataToWorkbook(sheetData: Sheet): XLSX.WorkBook {
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, userSettings.sheetName);

    return workbook;
}

function saveSheetFile(workbook: XLSX.WorkBook): void {
    XLSX.writeFile(workbook, `${userSettings.fileName}.${userSettings.fileExtension}`, { sheetStubs: true });
}
