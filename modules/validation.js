// Função para garantir que a string seja convertida em número e validada
function validateAndParseNumber(value) {
    if (typeof value === 'string') {
        value = parseFloat(value);
    }
    return typeof value === 'number' && !isNaN(value) ? value : null;
}

module.exports = {
    validateAndParseNumber
};
