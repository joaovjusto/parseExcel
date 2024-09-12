// Função para remover imagens da planilha
function removeImages(sheet) {
    if (sheet) {
        const images = sheet.getImages();
        images.forEach(img => {
            // Use o ID da imagem para remover
            sheet.workbook.model.media = sheet.workbook.model.media.filter(media => media.index !== img.imageId);
        });
        console.log(`Imagens removidas da planilha ${sheet.name}`);
    }
}

// Função para desbloquear todas as células de uma planilha
function unlockCells(sheet) {
    sheet.eachRow((row) => {
        row.eachCell((cell) => {
            cell.protection = { locked: false, hidden: false };
        });
    });
    console.log(`Todas as células da planilha ${sheet.name} foram desbloqueadas.`);
}

// Função para desbloquear todas as células de uma planilha e remover imagens
function unlockCellsAndRemoveImages(sheet) {
    unlockCells(sheet); // Chama a função para desbloquear células
    removeImages(sheet); // Chama a função para remover imagens
    console.log(`Todas as células da planilha ${sheet.name} foram desbloqueadas e imagens removidas.`);
}

// Função para modificar células com base na operação
async function modifySheets(workbook, sheetNames, workFunction) {
    const firstSheetName = workbook.worksheets[0].name;

    for (const sheetName of sheetNames) {
        const sheet = workbook.getWorksheet(sheetName);
        if (sheet) {
            await sheet.unprotect();
            if (sheetName === 'MENU') {
                unlockCellsAndRemoveImages(sheet); // Remove imagens da planilha MENU
            } else {
                unlockCells(sheet); // Desbloqueia as células das demais planilhas
            }

            if (sheetName !== 'MENU') {
                const linkCell = sheet.getCell('D7');
                linkCell.value = { text: "VOLTAR", hyperlink: `#'${firstSheetName}'!A1` };
                linkCell.style = { font: { bold: true, color: { argb: 'ffffffff' }, underline: false } };
                linkCell.protection = { locked: false, hidden: false };
                linkCell.alignment = { vertical: 'middle', horizontal: 'center' };
                linkCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D91919' } };
                linkCell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

                let rowIndex = 11;
                const column = 'C';

                while (true) {
                    const cellAddress = `${column}${rowIndex}`;
                    const cell = sheet.getCell(cellAddress);

                    if (typeof cell.value === 'number') {
                        let result = cell.value;
                        workFunction.forEach(op => {
                            result = `${result} ${op}`;
                        });
                        cell.value = eval(result);

                        console.log(`Célula ${cellAddress} modificada para: ${cell.value}`);
                    } else {
                        break;
                    }

                    rowIndex++;
                }
            }
        }
    }
}

module.exports = {
    unlockCellsAndRemoveImages,
    modifySheets
};
