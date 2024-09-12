const fs = require('fs');

// Função para verificar se o arquivo existe
function fileExists(filePath) {
    return fs.existsSync(filePath);
}

module.exports = {
    fileExists
};
