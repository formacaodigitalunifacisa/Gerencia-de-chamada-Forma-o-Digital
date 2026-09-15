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

    const data = agora.toLocaleDateString("pt-BR", {
        timeZone: "America/Sao_Paulo"
    });

    const hora = agora.toLocaleTimeString("pt-BR", {
        timeZone: "America/Sao_Paulo",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit"
    });


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