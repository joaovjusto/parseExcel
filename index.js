const { fileExists } = require('./modules/fileOperations');
const { askQuestion, closeReadline } = require('./modules/userInteraction');
const { validateAndParseNumber } = require('./modules/validation');
const { modifySheets } = require('./modules/excelOperations');
const ExcelJS = require('exceljs');

let workFunction = [];

async function modifyExcel() {
    let inputFileName;

    // Pergunta pelo nome do arquivo até encontrar um válido
    while (true) {
        inputFileName = await askQuestion('Digite o nome do arquivo de entrada (sem extensão .xlsx): ');
        if (fileExists(`./${inputFileName}.xlsx`)) {
            console.log('Arquivo encontrado:', `${inputFileName}.xlsx`);
            break;
        } else {
            console.log('Arquivo não encontrado. Tente novamente.');
        }
    }

    const outputFileName = await askQuestion('Digite o nome desejado para o arquivo de saída (sem extensão .xlsx): ');

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(`./${inputFileName}.xlsx`);

    const sheetNames = workbook.worksheets.map(sheet => sheet.name);
    const menuSheet = workbook.getWorksheet('MENU');
    if (menuSheet) {
        console.log('Desbloqueando células da planilha MENU');
        menuSheet.eachRow((row) => {
            row.eachCell((cell) => {
                cell.protection = { locked: false, hidden: false };
            });
        });
    }

    let actualQuestion = 0;
    const question = ['Digite o operador desejado (+, -, *, /) ou "sair" para encerrar: ', 'Digite o número desejado: '];

    while (true) {
        const operation = await askQuestion(question[actualQuestion]);

        if (actualQuestion === 0) {
            actualQuestion = 1;
        } else {
            actualQuestion = 0;
        }

        if (operation === 'sair') {
            console.log('Aplicando operações...');
            await modifySheets(workbook, sheetNames, workFunction);
            await workbook.xlsx.writeFile(`./${outputFileName}.xlsx`);
            console.log(`Arquivo salvo como "${outputFileName}.xlsx"`);
            break;
        } else if (validateAndParseNumber(operation)) {
            workFunction.push(operation);
        } else if (!['+', '-', '*', '/'].includes(operation)) {
            console.log('Operador inválido. Tente novamente.');
        } else {
            workFunction.push(operation);
        }
    }

    closeReadline();
}

modifyExcel();
