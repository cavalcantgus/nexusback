const { extractColumns } = require('./extract.js');
const { extractColumns1 } = require('./extract1.js');
const formatSheet1 = require('./format1.js');
const formatSheet = require('./format.js');
const XLSX = require('xlsx-js-style');

async function processExcelFile(inputFile, outputFile, reportType) {
    switch (reportType) {
        case 'Contas pagas por conta corrente':
            await contasPagasPorContaCorrente(inputFile, outputFile);
            break;
        case 'Títulos por data':
            await titulosPorData(inputFile, outputFile);
            break;
        default:
            console.log('Formatação não disponível para este tipo de relatório')
    }
}

async function titulosPorData(inputFile, outputFile) {
    try {
        const workbook = XLSX.readFile(inputFile, { cellStyles: true });
        const sheetName = workbook.SheetNames[0];
        let worksheet = workbook.Sheets[sheetName];

        const { newWorkSheet, centroCustoValues } = extractColumns1(worksheet);

        const formattedWorksheet = formatSheet1(newWorkSheet, centroCustoValues);

        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, formattedWorksheet, sheetName);

        // Salvar o arquivo formatado
        XLSX.writeFile(newWorkbook, outputFile, { bookType: 'xlsx', type: 'binary' });
        console.log('Arquivo processado e formatado com sucesso!');

        return newWorkbook;
    } catch (error) {
        console.error('Erro ao processar o arquivo:', error);
    }
}


async function contasPagasPorContaCorrente(inputFile, outputFile) {
    try {
        const workbook = XLSX.readFile(inputFile, { cellStyles: true });
        const sheetName = workbook.SheetNames[0];
        let worksheet = workbook.Sheets[sheetName];

        const { newWorkSheet, contaCorrenteValues } = extractColumns(worksheet);

        const formattedWorksheet = formatSheet(newWorkSheet, contaCorrenteValues);

        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, formattedWorksheet, sheetName);

        // Salvar o arquivo formatado
        XLSX.writeFile(newWorkbook, outputFile, { bookType: 'xlsx', type: 'binary' });
        console.log('Arquivo processado e formatado com sucesso!');

        return newWorkbook;
    } catch (error) {
        console.error('Erro ao processar o arquivo:', error);
    }
}

// Exportando a função
module.exports = {
    processExcelFile
};
