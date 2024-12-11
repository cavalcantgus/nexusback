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
    5,
    16, 17,       // QR
    24, 25,       // YZ
    26, 27, 28,   // AAABAC
    29, 30, 31,   // ADAEAF
    32, 33, 34, 35, // AGAHAIAJ
    36, 37, 38, 39  // AKALAMAN
];

const headersToRemove = [
    'Total da empresa',
    'Total do grupo de empresa',
    'Total geral',
    'SIENGE / SOFTPLAN'
];

function extractColumns(worksheet) {
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    const contaCorrenteValues = [];

    for(let rowIndex = 0; rowIndex < sheetData.length; rowIndex++) {
        const row = sheetData[rowIndex];
        const contaCorrenteIndex = row.slice(0, 5).indexOf('Conta Corrente');

        if(contaCorrenteIndex !== -1) {
            const value = row[contaCorrenteIndex + 5];
            contaCorrenteValues.push(value);
        }
    }

    const dataWithoutRows = removeRowsByHeaders(sheetData, headersToRemove);
    const updatedData = removeColumnsByIndexes(dataWithoutRows, columnsToRemove);

    const newWorkSheet = XLSX.utils.aoa_to_sheet(updatedData);

    delete newWorkSheet['!rows'];
    
    return { newWorkSheet, contaCorrenteValues };
}

// Exportando a função
module.exports = {
    extractColumns
};
