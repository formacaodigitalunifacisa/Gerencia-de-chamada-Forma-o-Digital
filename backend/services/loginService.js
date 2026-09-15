import { sheets, spreadsheetId } from "../config/googleSheets.js";


// ==========================================
// REALIZAR LOGIN
// ==========================================

async function realizarLogin(email, senha) {

    if (!email || !senha) {
        throw new Error("E-mail e senha são obrigatórios.");
    }


    // ==========================================
    // SENHA PADRÃO
    // ==========================================

    const senhaAdmin = process.env.ADMIN_PASSWORD;

    if (!senhaAdmin) {
        throw new Error("Senha do administrador não configurada.");
    }


    // ==========================================
    // VERIFICAR SENHA
    // ==========================================

    if (senha !== senhaAdmin) {
        throw new Error("Senha incorreta.");
    }


    // ==========================================
    // E-MAILS AUTORIZADOS
    // ==========================================

    const emailsAutorizados =
        process.env.ADMIN_EMAILS
            ?.split(",")
            .map(email => email.trim().toLowerCase())
            .filter(Boolean);


    if (!emailsAutorizados || emailsAutorizados.length === 0) {
        throw new Error("Nenhum administrador autorizado configurado.");
    }


    const emailNormalizado = email.trim().toLowerCase();


    if (!emailsAutorizados.includes(emailNormalizado)) {
        throw new Error("E-mail não autorizado.");
    }


    // ==========================================
    // BUSCAR ADMINISTRADOR NA PLANILHA
    // ==========================================

    const resposta = await sheets.spreadsheets.values.get({
        spreadsheetId: spreadsheetId,
        range: "Administradores!A2:C"
    });


    const administradores = resposta.data.values || [];


    // ==========================================
    // PROCURAR ADMINISTRADOR
    // ==========================================

    let administradorEncontrado = null;


    for (const linha of administradores) {

        const idAdministrador = linha[0];
        const nome = linha[1];
        const emailAdministrador =
            linha[2]?.trim().toLowerCase();


        if (emailAdministrador === emailNormalizado) {

            administradorEncontrado = {
                idAdministrador: idAdministrador,
                nome: nome,
                email: emailAdministrador
            };

            break;
        }
    }


    // ==========================================
    // ADMINISTRADOR NÃO ENCONTRADO
    // ==========================================

    if (!administradorEncontrado) {
        throw new Error(
            "Administrador não encontrado na planilha."
        );
    }


    // ==========================================
    // LOGIN REALIZADO
    // ==========================================

    return administradorEncontrado;
}


export {
    realizarLogin
};