import { sheets, spreadsheetId } from "../config/googleSheets.js";
import { registrarAlteracao } from "./logService.js";


// ==========================================
// ADICIONAR TURMA
// ==========================================

async function adicionarTurma(turma, administrador) {

    if (!turma) {
        throw new Error("Número da turma é obrigatório.");
    }


    const numeroTurma = String(turma).trim();


    if (!numeroTurma) {
        throw new Error("Número da turma é obrigatório.");
    }


    // ==========================================
    // VERIFICAR SE É UM NÚMERO
    // ==========================================

    if (!/^\d+$/.test(numeroTurma)) {
        throw new Error("O número da turma deve ser informado apenas com números.");
    }


    // ==========================================
    // BUSCAR AS ABAS EXISTENTES
    // ==========================================

    const planilha = await sheets.spreadsheets.get({
        spreadsheetId
    });

    const abas = planilha.data.sheets || [];


    const nomeAbaAlunos = `AlunosTurma${numeroTurma}`;
    const nomeAbaVoluntarios = `VoluntariosTurma${numeroTurma}`;


    // ==========================================
    // VERIFICAR SE A TURMA JÁ EXISTE
    // ==========================================

    const alunosExiste = abas.some(
        aba => aba.properties.title === nomeAbaAlunos
    );

    const voluntariosExiste = abas.some(
        aba => aba.properties.title === nomeAbaVoluntarios
    );


    if (alunosExiste || voluntariosExiste) {
        throw new Error("Essa turma já existe.");
    }


    // ==========================================
    // CRIAR AS DUAS ABAS
    // ==========================================

    await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
            requests: [
                {
                    addSheet: {
                        properties: {
                            title: nomeAbaAlunos
                        }
                    }
                },
                {
                    addSheet: {
                        properties: {
                            title: nomeAbaVoluntarios
                        }
                    }
                }
            ]
        }
    });


    // ==========================================
    // CRIAR CABEÇALHO DOS ALUNOS
    // ==========================================

    await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${nomeAbaAlunos}!A1:B1`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
            values: [
                [
                    "ID",
                    "Nome"
                ]
            ]
        }
    });


    // ==========================================
    // CRIAR CABEÇALHO DOS VOLUNTÁRIOS
    // ==========================================

    await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${nomeAbaVoluntarios}!A1:B1`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
            values: [
                [
                    "ID",
                    "Nome"
                ]
            ]
        }
    });


    // ==========================================
    // REGISTRAR NO LOG
    // ==========================================

    await registrarAlteracao(
        administrador.idAdministrador,
        administrador.nome,
        administrador.email,
        nomeAbaAlunos,
        "ADICIONAR",
        `Turma ${numeroTurma} criada com as abas ${nomeAbaAlunos} e ${nomeAbaVoluntarios}.`
    );


    // ==========================================
    // RETORNAR RESULTADO
    // ==========================================

    return {
        turma: numeroTurma,
        alunos: nomeAbaAlunos,
        voluntarios: nomeAbaVoluntarios,
        mensagem: "Turma adicionada com sucesso."
    };
}


// ==========================================
// LISTAR TURMAS
// ==========================================

async function listarTurmas() {

    const planilha = await sheets.spreadsheets.get({
        spreadsheetId
    });

    const abas = planilha.data.sheets || [];

    const turmas = [];


    for (const aba of abas) {

        const nomeAba =
            aba.properties.title;


        if (nomeAba.startsWith("AlunosTurma")) {

            const numeroTurma =
                nomeAba.replace("AlunosTurma", "");


            const voluntariosExiste =
                abas.some(
                    outraAba =>
                        outraAba.properties.title ===
                        `VoluntariosTurma${numeroTurma}`
                );


            if (voluntariosExiste) {

                turmas.push({
                    numero: numeroTurma,
                    nome: `Turma ${numeroTurma}`,
                    alunos: `AlunosTurma${numeroTurma}`,
                    voluntarios: `VoluntariosTurma${numeroTurma}`
                });

            }

        }

    }


    turmas.sort((a, b) =>
        Number(a.numero) - Number(b.numero)
    );


    return turmas;
}


export {
    adicionarTurma,
    listarTurmas
};