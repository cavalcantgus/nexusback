const XLSX = require('xlsx-js-style');

const headersFont = {
    name: 'Red Hat',
    sz: 10,
    color: { rgb: "FFFFFF" },
    bold: false,
    italic: false
}

const redHatFont = {
    name: 'Red Hat',
    sz: 10,
    color: { rgb: "000000" },
    bold: false,
    italic: false
};

const titleFont = {
    name: 'Red Hat',
    sz: 24,
    color: { rgb: "000000" },
    bold: true,
    italic: false
};

const headersToFormat = [
    'Período de emissão',
    'Período de cadastro',
    'Empresa',
    'Centro de custo',
    'Item',
    'Titulo',
    'Credor',
    'Documento',
    'Emis.NF',
    'Cadastro',
    'Vencto',
    'Qtd',
    'Valor líquido',
    'Total centro custo',
    'Total empresa',
    'Total geral'
];

const valuesToMerge = [
    'Período de cadastro',
    'Empresa',
    'Centro de custo'
];

const headersToMerge = [
    'Total centro custo',
    'Total empresa'
];

const currencyFormat = {
    numFmt: 'R$ #,##0.00' // Formato de moeda contábil
};

const targetColumnEnd = 8;

const headerRow = 12

function formatSheet1(worksheet, contaCorrenteValues) {
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (!worksheet['!ref']) {
        throw new Error('A referência da planilha está indefinida.');
    }

    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const merges = [];
    let contaCorrenteIndex = 0;

    const defaultBorder = {
        top: { style: "thin", color: { auto: 1 } },
        bottom: { style: "thin", color: { auto: 1 } },
        left: { style: "thin", color: { auto: 1 } },
        right: { style: "thin", color: { auto: 1 } }
    };

    const titleBorder = {
        bottom: { style: "dotted", color: { auto: 1 }},
        right: { style: "dotted", color: { auto: 1 }}
    };

    const titleCellAddress = XLSX.utils.encode_cell({ r: 1, c: 0 });
    const titleValue = worksheet[titleCellAddress] ? worksheet[titleCellAddress].v : "Título do Relatório";
    
    // Mesclando três linhas (1, 2 e 3) nas colunas A até F
    for (let row = 0; row < 3; row++) {
        const titleMergeStart = XLSX.utils.encode_cell({ r: row, c: 0 });
        const titleMergeEnd = XLSX.utils.encode_cell({ r: row, c: 8});
        merges.push({
            s: XLSX.utils.decode_cell(titleMergeStart),
            e: XLSX.utils.decode_cell(titleMergeEnd)
        });

        // Aplicando a formatação de fundo nas células mescladas
        for (let col = 0; col <= 5; col++) { // Colunas de A (0) até F (5)
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            worksheet[cellAddress] = {
                s: {
                    fill: { fgColor: { rgb: "FFFFFF" } }, // Cor de fundo azul
                }
            };
        }

        if (row === 1) {
            worksheet[titleMergeStart] = {
                v: titleValue,
                s: {
                    fill: { fgColor: { rgb: "FFFFFF" } },
                    font: titleFont,
                    alignment: { horizontal: "center", vertical: "center" },
                }
            };
        } else {
            worksheet[titleMergeStart] = {
                s: {
                    fill: { fgColor: { rgb: "FFFFFF" } },
                    border: titleBorder
                }
            };
        }
    }

    for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex++) {
        const row = sheetData[rowIndex] || [];

        for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex++) {
            const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
            const cellValue = worksheet[cellAddress] ? worksheet[cellAddress].v : null;

            if (!worksheet[cellAddress]) {
                worksheet[cellAddress] = {};
            }

            // Limpar a formatação se a célula for parte do título
            if (rowIndex < 3) {
                continue; // Pular formatação para as linhas do título
            }

            // Aplica estilo para cabeçalhos formatados
            if (headersToFormat.includes(cellValue)) {
                worksheet[cellAddress].s = {
                    fill: { fgColor: { rgb: "01445f" } },
                    font: headersFont,
                    alignment: { horizontal: "left" },
                    border: defaultBorder
                };
            } else {
                worksheet[cellAddress].s = { border: defaultBorder }; // Aplicar bordas
            }

            if (headersToMerge.includes(cellValue)) {
                const startAddress = XLSX.utils.encode_cell({ r: rowIndex, c: 0 });
                const endAddress = XLSX.utils.encode_cell({ r: rowIndex, c: 7 });

                merges.push({
                    s: XLSX.utils.decode_cell(startAddress),
                    e: XLSX.utils.decode_cell(endAddress)
                });

                const valueAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex + 1 });
                let headerValue = worksheet[valueAddress] ? worksheet[valueAddress].v : null;

                // Se o valor não for nulo, aplica a conversão para float
                if (headerValue !== null) {
                    // Remove pontos usados como separador de milhar
                    let sanitizedValue = String(headerValue).replace(/\.(?=\d{3})/g, '');

                    // Substitui a vírgula decimal por um ponto para conversão em float
                    sanitizedValue = sanitizedValue.replace(',', '.');

                    // Converte para float
                    const floatValue = parseFloat(sanitizedValue);

                    // Verifica se a conversão foi bem-sucedida
                    if (!isNaN(floatValue)) {
                        // Atualiza a célula com o valor numérico (não formatado como string)
                        worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: 8 })] = {
                            v: floatValue,  // Valor numérico puro
                            t: 'n',         // Indica que é um número
                            s: {
                                font: redHatFont,
                                numFmt: '"R$" #,##0.00_);[Red]("R$" #,##0.00)',  // Formato numérico de moeda
                                alignment: { horizontal: "right" },
                                border: defaultBorder,
                            }
                        };
                    } else {
                        console.log(`Falha ao converter o valor: ${headerValue}`);
                    }
                }
            }

            
            
            // Formatação específica para "Conta Corrente"
            if (colIndex === 0 && cellValue === 'Centro de custo') {
                if (rowIndex >= 10 && contaCorrenteValues[contaCorrenteIndex] !== undefined) {
                    const nextColumnAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex + 1 });

                    worksheet[nextColumnAddress] = {
                        v: contaCorrenteValues[contaCorrenteIndex],
                        s: {
                            alignment: { horizontal: "right" },
                            border: defaultBorder,
                        }
                    };

                    contaCorrenteIndex++;
                }
            }

            // Aplicar bordas e alinhar todas as células à direita, exceto as linhas do título
            if (rowIndex !== 1) {
                worksheet[cellAddress].s = {
                    ...worksheet[cellAddress].s,
                    alignment: {
                        horizontal: "right",
                        vertical: "top", // Alinha o texto ao topo para facilitar a leitura
                        wrapText: true   // Habilita a quebra de linha automática
                    },
                    border: defaultBorder
                };
            }
        }
    }

    for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex++) {
        const row = sheetData[rowIndex] || [];

        for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex++) {
            const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
            const cellValue = worksheet[cellAddress] ? worksheet[cellAddress].v : null;

            // Verifica se a célula contém algum dos headers a serem mesclados
            if (valuesToMerge.includes(cellValue)) {
                const nextColIndex = colIndex + 1; // Coluna à direita
                const valueAddress = XLSX.utils.encode_cell({ r: rowIndex, c: nextColIndex });
                const associatedValue = worksheet[valueAddress] ? worksheet[valueAddress].v : null;

                if (associatedValue !== null && associatedValue !== undefined) {
                    // Mesclar da coluna à direita até a coluna F (coluna 5)
                    const startAddress = XLSX.utils.encode_cell({ r: rowIndex, c: nextColIndex });
                    const endAddress = XLSX.utils.encode_cell({ r: rowIndex, c: targetColumnEnd });

                    merges.push({
                        s: XLSX.utils.decode_cell(startAddress),
                        e: XLSX.utils.decode_cell(endAddress),
                        border: defaultBorder
                    });

                    // Inserir o valor associado na célula mesclada
                    worksheet[startAddress] = {
                        v: associatedValue,
                        s: {
                            font: redHatFont,
                            alignment: { horizontal: "left", vertical: "center", wrapText: true },
                            border: defaultBorder
                        }
                    };

                    // Limpa as células mescladas (da coluna seguinte até F) para evitar valores residuais
                    for (let mergeColIndex = nextColIndex + 1; mergeColIndex <= targetColumnEnd; mergeColIndex++) {
                        const clearCellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: mergeColIndex });
                        worksheet[clearCellAddress] = { s: { border: defaultBorder } };
                    }
                }
            }
        }
    }

    function isColumnI(cellAddress) {
        return cellAddress.startsWith('I'); // Verifica se o endereço começa com "I"
    }

    for (let cellAddress in worksheet) {
        if (worksheet.hasOwnProperty(cellAddress)) {
            const cell = worksheet[cellAddress];
            
            // Se a célula está na coluna I e o valor é numérico, aplica o formato contábil
            if (isColumnI(cellAddress) && typeof cell.v === 'number') {
                worksheet[cellAddress].s = {
                    numFmt: '"R$" #,##0.00_);[Red]("R$" #,##0.00)',
                    font: redHatFont,
                    border: defaultBorder
                 }
            }
        }
    }
    
    worksheet['!merges'] = merges;

    worksheet['!cols'] = [
        { wpx: 200 },
        { wpx: 300 },
        { wpx: 200 },
        { wpx: 100 },
        { wpx: 100 },
        { wpx: 100 },
        { wpx: 100 },
        { wpx: 50},
        { wpx: 100}
    ];

    return worksheet;
}

// Exportando a função
module.exports = formatSheet1;
