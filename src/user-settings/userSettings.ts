export const userSettings = {
    fileName: 'Norakstit',
    fileExtension: 'xlsx',
    sheetName: 'Test',
    font: 'Times New Roman',
    ignoreEmptyTable: true,
    replaceStrings: [
        {
            searchValue: 'Valdes priekšsēdētājs',
            replaceValue: 'Noliktavas prečzinis',
        },
    ],
    applyPercentFormula: [
        {
            tableWithHeaders: [
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
            options: {
                percentCellHeaders: [
                    '1. Brokastis',
                    '2. Pusdienas',
                    '3. Launags',
                    '4. Vakariņas',
                ],
                totalCellHeader: 'Kopā',
            },
        },
    ],
};
