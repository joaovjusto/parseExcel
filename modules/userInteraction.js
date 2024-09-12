const readline = require('readline');

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

// Função para perguntar ao usuário sobre a operação
function askQuestion(query) {
    return new Promise(resolve => rl.question(query, resolve));
}

// Função para fechar a interface readline
function closeReadline() {
    rl.close();
}

module.exports = {
    askQuestion,
    closeReadline
};
