import { SheetCell } from '../../types';
import { userSettings } from '../../user-settings';

const getValue = (value: string): string | number => {
    let mutatedValue: string | number = value;

    userSettings.replaceStrings.forEach((item) => {
        if ((mutatedValue as String).includes(item.searchValue)) {
            mutatedValue = (mutatedValue as String).replace(item.searchValue, item.replaceValue);
        }
    });

    if (value.match(/^\d+,\d+$/) || value.match(/^\d+\.\d+$/) || value.match(/^\d+$/)) {
        if (value.match(/^\d+,\d+$/)) {
            mutatedValue = mutatedValue.replace(',', '.')
        }

        mutatedValue = Number(mutatedValue);
    }

    return mutatedValue;
};

const getCell = (v: string | number): SheetCell => {
    let s = {};

    // TODO: Copy styles
    if (typeof v === 'string' && v.match(/^Produktu pieprasījums/)) {
        s = {
            font: {
                bold: true,
            }
        }
    }

    return {
        t: typeof v === 'string' ? 's' : 'n',
        v,
        s,
    };
};

const createEmpty = (): SheetCell => {
    return {
        t: 's',
        v: ''
    };
};

export const cell = { getValue, getCell, createEmpty };
