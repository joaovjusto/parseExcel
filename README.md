
# Excel Modifier Project

Este projeto permite modificar arquivos Excel de maneira interativa utilizando ExcelJS e readline. Ele desbloqueia as células de planilhas, aplica operações em células numéricas e remove todas as imagens da planilha chamada "MENU".

## Funcionalidades
- Desbloqueia todas as células de planilhas no arquivo Excel.
- Remove todas as imagens da planilha "MENU".
- Permite ao usuário aplicar operações matemáticas (+, -, *, /) em células da coluna "C" de outras planilhas.
- Adiciona um hiperlink na célula "D7" para voltar à primeira planilha do arquivo.
- Salva o arquivo modificado com um novo nome fornecido pelo usuário.

## Dependências
- Node.js
- ExcelJS
- Readline (nativo no Node.js)
- File System (nativo no Node.js)

## Instalação

Clone o repositório ou extraia os arquivos.
Instale as dependências do projeto executando:

```bash
npm install
```

## Uso

Execute o script principal:

```bash
node index.js
```

O script solicitará:

- O nome do arquivo de entrada (sem a extensão .xlsx).
- O nome desejado para o arquivo de saída.
- A operação matemática a ser aplicada às células (+, -, *, / ou "sair" para finalizar).
- Os valores a serem aplicados nas operações.

O arquivo modificado será salvo com o nome fornecido.

## Estrutura do Projeto

```css
.
├── node_modules
├── src
│   ├── excelOperations.js
│   └── main.js
├── index.js
├── package.json
└── README.md
```

- `index.js`: Arquivo principal que inicia o script.
- `src/excelOperations.js`: Contém as funções para manipulação do arquivo Excel.
- `src/main.js`: Controla o fluxo principal da aplicação, incluindo as interações com o usuário.
- `package.json`: Contém as dependências do projeto.

## Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir um pull request ou reportar problemas.

## Licença

Este projeto está sob a licença MIT.
