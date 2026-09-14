import { sheets, spreadsheetId } from "./config/googleSheets.js";

async function testarConexao() {
    try {
        const resposta = await sheets.spreadsheets.get({
            spreadsheetId: spreadsheetId
        });

        console.log("Conexão realizada com sucesso!");
        console.log("Planilha:", resposta.data.properties.title);

    } catch (erro) {
        console.error("Erro ao conectar com o Google Sheets:");
        console.error(erro.message);
    }
}

testarConexao();