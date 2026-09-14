import { sheets, spreadsheetId } from "../config/googleSheets.js";


async function realizarLogin(email, senha) {

    if (!email || !senha) {
        throw new Error("E-mail e senha são obrigatórios.");
    }


    // Busca os administradores cadastrados
    const resposta = await sheets.spreadsheets.values.get({
        spreadsheetId: spreadsheetId,
        range: "Administradores!A2:D"
    });


    const administradores = resposta.data.values || [];


    // Procura o e-mail
    let administradorEncontrado = null;

    for (const linha of administradores) {

        const idAdministrador = linha[0];
        const nome = linha[1];
        const emailAdministrador = linha[2];
        const senhaAdministrador = linha[3];


        if (emailAdministrador === email) {

            administradorEncontrado = {
                idAdministrador: idAdministrador,
                nome: nome,
                email: emailAdministrador,
                senha: senhaAdministrador
            };

            break;
        }
    }


    // E-mail não cadastrado
    if (!administradorEncontrado) {
        throw new Error("E-mail não autorizado.");
    }


    // Senha incorreta
    if (administradorEncontrado.senha !== senha) {
        throw new Error("Senha incorreta.");
    }


    // Login realizado
    return {
        idAdministrador: administradorEncontrado.idAdministrador,
        nome: administradorEncontrado.nome,
        email: administradorEncontrado.email
    };
}


export {
    realizarLogin
};