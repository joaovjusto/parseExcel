const ExcelJS = require('exceljs');
const readline = require('readline');
const fs = require('fs');

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

// Função para garantir que a string seja convertida em número e validada
function validateAndParseNumber(value) {
    if (typeof value === 'string') {
        value = parseFloat(value);
    }
    return typeof value === 'number' && !isNaN(value) ? value : null;
}

// Função para perguntar ao usuário sobre a operação
function askQuestion(query) {
    return new Promise(resolve => rl.question(query, resolve));
}

let workFunction = [];

// Função para verificar se o arquivo existe
function fileExists(filePath) {
    return fs.existsSync(filePath);
}

// Função principal para modificar o Excel
async function modifyExcel() {
    let inputFileName;

    // Pergunta ao usuário pelo nome do arquivo de entrada até que um arquivo válido seja fornecido
    while (true) {
        inputFileName = await askQuestion('Digite o nome do arquivo de entrada (sem extensão .xlsx): ');

        // Verifique se o arquivo existe
        if (fileExists(`./${inputFileName}.xlsx`)) {
            console.log('Arquivo encontrado:', `${inputFileName}.xlsx`);
            break;
        } else {
            console.log('Arquivo não encontrado. Por favor, tente novamente.');
        }
    }

    const outputFileName = await askQuestion('Digite o nome desejado para o arquivo de saída (sem extensão .xlsx): ');

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(`./${inputFileName}.xlsx`);

    // Liste os nomes das planilhas disponíveis
    const sheetNames = workbook.worksheets.map(sheet => sheet.name);

    // Desbloquear todas as células da planilha MENU
    const menuSheet = workbook.getWorksheet('MENU');
    if (menuSheet) {
        console.log('Desbloqueando todas as células da planilha: MENU');
        menuSheet.eachRow((row, rowIndex) => {
            row.eachCell((cell, colNumber) => {
                cell.protection = { locked: false, hidden: false };
            });
        });
        console.log('Todas as células da planilha MENU foram desbloqueadas.');
    }

    // Resto do código...
    let operation;
    let actualQuestion = 0;
    let question = ['Digite o operador desejado (+, -, *, /) ou "sair" para encerrar: ', 'Digite o número desejado: '];

    // Pergunta ao usuário qual operação deseja realizar
    while (true) {
        operation = await askQuestion(question[actualQuestion]);

        if (actualQuestion === 0) {
            actualQuestion = 1;
        } else {
            actualQuestion = 0;
        }

        // Valide a operação
        if (operation === 'sair') {
            console.log('Aplicando operação:', workFunction);

            const firstSheetName = workbook.worksheets[0].name; // Nome da primeira planilha

            for (const sheetName of sheetNames) {
                if (sheetName !== 'MENU') {
                    console.log('Modificando a planilha:', sheetName);
                    const sheet = workbook.getWorksheet(sheetName); // Acessa a planilha pelo nome

                    if (sheet) {
                        console.log('Planilha carregada:', sheet.name);

                        // Remova a proteção da planilha se houver
                        if (sheet.protection) {
                            await sheet.unprotect();
                            console.log(`Proteção removida da planilha: ${sheet.name}`);
                        }

                        // Desbloqueie e exiba todas as células
                        sheet.eachRow((row, rowIndex) => {
                            row.eachCell((cell, colNumber) => {
                                cell.protection = { locked: false, hidden: false };
                            });
                        });

                        // Adicionar hiperlink na célula D7
                        const linkCell = sheet.getCell('D7');
                        linkCell.value = { text: "VOLTAR", hyperlink: `#'${firstSheetName}'!A1` };
                        linkCell.style = { font: { bold: true, color: { argb: 'ffffffff' }, underline: false, } };
                        linkCell.protection = { locked: false, hidden: false };
                        linkCell.alignment = { vertical: 'middle', horizontal: 'center' };
                        linkCell.fill = {
                            type: 'pattern',
                            pattern:'solid',
                            fgColor:{argb:'D91919'},
                          };
                        linkCell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

                        console.log(`Hiperlink adicionado em ${sheet.name} na célula E7 para ${firstSheetName}.`);

                        let rowIndex = 11; // Começa na linha 11
                        const column = 'C'; // Coluna C

                        while (true) {
                            const cellAddress = `${column}${rowIndex}`;
                            const cell = sheet.getCell(cellAddress);

                            if (typeof cell.value === 'number') {
                                let oldValue = cell.value;
                                let result = cell.value;
                                workFunction.forEach((value) => {
                                    result = `${result} ${value}`;
                                });
                                cell.value = eval(result);

                                console.log(`Célula de valor ${oldValue} modificada para: ${cell.value}`);
                            } else {
                                // Quando encontramos uma célula que não é um número, saímos do loop
                                console.log(`Célula ${cellAddress} não contém um número ou não existe mais.`);
                                break;
                            }

                            rowIndex++; // Avance para a próxima linha
                        }
                    } else {
                        console.log('Planilha não encontrada.');
                    }
                }
            }


            // Salve as alterações em um novo arquivo
            await workbook.xlsx.writeFile(`./${outputFileName}.xlsx`);
            console.log(`Arquivo salvo com sucesso como "${outputFileName}.xlsx".`);
            console.log('Saindo...');
            break;
        } else if (validateAndParseNumber(operation)) {
            workFunction.push(operation);
        } else if (!['+', '-', '*', '/'].includes(operation)) {
            console.log('Operador inválido. Tente novamente.');
        } else {
            workFunction.push(operation);
        }
    }

    rl.close();
}

modifyExcel();