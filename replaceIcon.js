const rcedit = require('rcedit');
const path = require('path');

const exePath = path.join(__dirname, 'dist', 'parseExcel-win.exe');
const iconPath = path.join(__dirname, 'foto.ico');

rcedit(exePath, {
    'icon': iconPath
}, function(err) {
    if (err) {
        console.error('Erro ao substituir o ícone:', err);
    } else {
        console.log('Ícone substituído com sucesso!');
    }
});
