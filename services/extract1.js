const XLSX = require('xlsx-js-style');

function removeColumnsByIndexes(sheetData, columnsToRemove) {
    return sheetData.map(row => 
        row.filter((_, index) => !columnsToRemove.includes(index))
    );
}

function removeRowsByHeaders(sheetData, headersToRemove) {
    return sheetData.filter(row =>
        !headersToRemove.some(header => row.includes(header))
    );
}

const columnsToRemove = [
    7, 8, 9,
    10, 11,       // QR
    14, 13, // AAABAC
    17, 18,   // ADAEAF
    22, 23, 24, 25, // AGAHAIAJ
    26
];

const headersToRemove = [
    'Total geral'
];

function extractColumns1(worksheet) {
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    const centroCustoValues = [];

    for(let rowIndex = 0; rowIndex < sheetData.length; rowIndex++) {
        const row = sheetData[rowIndex];
        const centroCustoIndex = row.slice(0, 5).indexOf('Centro de custo');

        if(centroCustoIndex !== -1) {
            const value = row[centroCustoIndex + 5];
            centroCustoValues.push(value);
        }
    }

    const dataWithoutRows = removeRowsByHeaders(sheetData, headersToRemove);
    const updatedData = removeColumnsByIndexes(dataWithoutRows, columnsToRemove);

    const newWorkSheet = XLSX.utils.aoa_to_sheet(updatedData);

    delete newWorkSheet['!rows'];
    
    return { newWorkSheet, centroCustoValues };
}

// Exportando a função
module.exports = {
    extractColumns1
};
