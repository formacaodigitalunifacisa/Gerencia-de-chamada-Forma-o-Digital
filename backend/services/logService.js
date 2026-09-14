import { sheets, spreadsheetId } from "../config/googleSheets.js";


// ==========================================
// REGISTRAR ALTERAÇÃO
// ==========================================

async function registrarAlteracao(
    idAdministrador,
    nomeAdministrador,
    emailAdministrador,
    aba,
    tipo,
    descricao
) {

    const agora = new Date();

    const data = agora.toLocaleDateString("pt-BR");
    const hora = agora.toLocaleTimeString("pt-BR");


    await sheets.spreadsheets.values.append({
        spreadsheetId: spreadsheetId,
        range: "LogAlteracoes!A:I",
        valueInputOption: "USER_ENTERED",
        insertDataOption: "INSERT_ROWS",
        requestBody: {
            values: [
                [
                    null,
                    idAdministrador,
                    nomeAdministrador,
                    emailAdministrador,
                    data,
                    hora,
                    aba,
                    tipo,
                    descricao
                ]
            ]
        }
    });

}


export {
    registrarAlteracao
};