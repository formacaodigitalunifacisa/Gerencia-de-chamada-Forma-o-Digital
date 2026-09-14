import { sheets, spreadsheetId } from "../config/googleSheets.js";
import { registrarAlteracao } from "./logService.js";


// ==========================================
// CONVERTER NÚMERO PARA LETRA DA COLUNA
// ==========================================

function numeroParaColuna(numero) {

    let coluna = "";

    while (numero > 0) {

        const resto =
            (numero - 1) % 26;

        coluna =
            String.fromCharCode(
                65 + resto
            ) + coluna;

        numero =
            Math.floor(
                (numero - 1) / 26
            );
    }

    return coluna;
}


// ==========================================
// CONVERTER TURMA PARA NOME DA ABA
// ==========================================

function obterNomeAbaVoluntarios(turma) {

    const valor =
        String(turma).trim();


    if (
        valor.startsWith("VoluntariosTurma")
    ) {

        return valor;
    }


    if (
        valor.startsWith("Turma ")
    ) {

        const numero =
            valor.replace("Turma ", "").trim();

        return `VoluntariosTurma${numero}`;
    }


    return `VoluntariosTurma${valor}`;
}


// ==========================================
// LIMPAR LINHA DE CHAMADA
// ==========================================

async function limparLinhaChamada(
    sheetId,
    nomeAba,
    numeroLinha
) {

    const primeiraLinha =
        await sheets.spreadsheets.values.get({

            spreadsheetId:
            spreadsheetId,

            range:
                `${nomeAba}!1:1`
        });


    const cabecalho =
        primeiraLinha.data.values?.[0] || [];


    if (cabecalho.length <= 2) {
        return;
    }


    const ultimaColuna =
        numeroParaColuna(
            cabecalho.length
        );


    await sheets.spreadsheets.values.clear({

        spreadsheetId:
        spreadsheetId,

        range:
            `${nomeAba}!C${numeroLinha}:${ultimaColuna}${numeroLinha}`
    });


    await sheets.spreadsheets.batchUpdate({

        spreadsheetId:
        spreadsheetId,

        requestBody: {

            requests: [

                {

                    repeatCell: {

                        range: {

                            sheetId:
                            sheetId,

                            startRowIndex:
                                numeroLinha - 1,

                            endRowIndex:
                            numeroLinha,

                            startColumnIndex:
                                2,

                            endColumnIndex:
                            cabecalho.length
                        },

                        cell: {

                            userEnteredFormat: {

                                backgroundColor: {

                                    red: 1,
                                    green: 1,
                                    blue: 1
                                },

                                horizontalAlignment:
                                    "CENTER",

                                verticalAlignment:
                                    "MIDDLE"
                            }
                        },

                        fields:
                            "userEnteredFormat.backgroundColor," +
                            "userEnteredFormat.horizontalAlignment," +
                            "userEnteredFormat.verticalAlignment"
                    }
                }
            ]
        }
    });
}


// ==========================================
// BUSCAR PRÓXIMO ID DO VOLUNTÁRIO
// ==========================================

async function obterProximoIdVoluntario(
    planilha
) {

    let maiorId = 0;


    // ==========================================
    // 1. HISTÓRICO
    // ==========================================

    try {

        const historico =
            await sheets.spreadsheets.values.get({

                spreadsheetId:
                spreadsheetId,

                range:
                    "HistoricoVoluntario!A2:A"
            });


        const linhasHistorico =
            historico.data.values || [];


        for (
            const linha of linhasHistorico
            ) {

            const id =
                Number(linha[0]);


            if (
                Number.isInteger(id) &&
                id > maiorId
            ) {

                maiorId = id;
            }
        }

    } catch (erro) {

        console.warn(
            "Não foi possível consultar o histórico dos voluntários:",
            erro.message
        );
    }


    // ==========================================
    // 2. ABAS DOS VOLUNTÁRIOS
    // ==========================================

    for (
        const aba of
    planilha.data.sheets || []
        ) {

        const nomeAba =
            aba.properties.title;


        if (
            !nomeAba.startsWith(
                "VoluntariosTurma"
            )
        ) {

            continue;
        }


        try {

            const resposta =
                await sheets.spreadsheets.values.get({

                    spreadsheetId:
                    spreadsheetId,

                    range:
                        `${nomeAba}!A2:A`
                });


            const linhas =
                resposta.data.values || [];


            for (
                const linha of linhas
                ) {

                const id =
                    Number(linha[0]);


                if (
                    Number.isInteger(id) &&
                    id > maiorId
                ) {

                    maiorId = id;
                }
            }

        } catch (erro) {

            console.warn(
                `Não foi possível consultar a aba ${nomeAba}:`,
                erro.message
            );
        }
    }


    // ==========================================
    // 3. LOGS
    // ==========================================

    try {

        const logs =
            await sheets.spreadsheets.values.get({

                spreadsheetId:
                spreadsheetId,

                range:
                    "LogAlteracoes!A2:I"
            });


        const linhasLogs =
            logs.data.values || [];


        for (
            const linha of linhasLogs
            ) {

            const descricao =
                String(linha[8] || "");


            const correspondencia =
                descricao.match(
                    /\bVolunt(?:ário|ario)\b.*?\bID\s*[:#-]?\s*(\d+)/i
                );


            if (correspondencia) {

                const id =
                    Number(
                        correspondencia[1]
                    );


                if (
                    Number.isInteger(id) &&
                    id > maiorId
                ) {

                    maiorId = id;
                }
            }
        }

    } catch (erro) {

        console.warn(
            "Não foi possível consultar LogAlteracoes para ID do voluntário:",
            erro.message
        );
    }


    return maiorId + 1;
}


// ==========================================
// ADICIONAR VOLUNTÁRIO
// ==========================================

async function adicionarVoluntario(
    nome,
    turma,
    curso,
    outroCurso,
    administrador
) {

    if (
        !nome ||
        !turma ||
        !curso
    ) {

        throw new Error(
            "Nome, turma e curso são obrigatórios."
        );
    }


    const nomeVoluntario =
        String(nome).trim();


    const turmaInformada =
        String(turma).trim();


    if (!nomeVoluntario) {

        throw new Error(
            "Nome do voluntário é obrigatório."
        );
    }


    if (!turmaInformada) {

        throw new Error(
            "Turma é obrigatória."
        );
    }


    const cursosPermitidos = [

        "Sistemas de Informação",
        "ADS",
        "Fisioterapia",
        "Enfermagem",
        "Medicina",
        "Outro"
    ];


    if (
        !cursosPermitidos.includes(
            curso
        )
    ) {

        throw new Error(
            "Curso inválido."
        );
    }


    if (
        curso === "Outro" &&
        !outroCurso
    ) {

        throw new Error(
            "Informe o nome do outro curso."
        );
    }


    let cursoFinal =
        curso;


    if (
        curso === "Outro"
    ) {

        cursoFinal =
            String(outroCurso).trim();


        if (!cursoFinal) {

            throw new Error(
                "Informe o nome do outro curso."
            );
        }
    }


    // ==========================================
    // BUSCAR PLANILHA
    // ==========================================

    const planilha =
        await sheets.spreadsheets.get({

            spreadsheetId:
            spreadsheetId
        });


    const nomeAbaTurma =
        obterNomeAbaVoluntarios(
            turmaInformada
        );


    // ==========================================
    // VERIFICAR TURMA
    // ==========================================

    const abaTurma =
        planilha.data.sheets.find(

            sheet =>
                sheet.properties.title ===
                nomeAbaTurma
        );


    if (!abaTurma) {

        throw new Error(
            "Turma não encontrada."
        );
    }


    // ==========================================
    // NOVO ID
    // ==========================================

    const novoId =
        await obterProximoIdVoluntario(
            planilha
        );


    // ==========================================
    // 1. HISTÓRICO PRIMEIRO
    // ==========================================

    await sheets.spreadsheets.values.append({

        spreadsheetId:
        spreadsheetId,

        range:
            "HistoricoVoluntario!A:F",

        valueInputOption:
            "USER_ENTERED",

        insertDataOption:
            "INSERT_ROWS",

        requestBody: {

            values: [

                [

                    novoId,

                    nomeVoluntario,

                    cursoFinal,

                    0,

                    0,

                    nomeAbaTurma
                ]
            ]
        }
    });


    // ==========================================
    // 2. ADICIONAR NA TURMA
    // ==========================================

    await sheets.spreadsheets.values.append({

        spreadsheetId:
        spreadsheetId,

        range:
            `${nomeAbaTurma}!A:B`,

        valueInputOption:
            "USER_ENTERED",

        insertDataOption:
            "INSERT_ROWS",

        requestBody: {

            values: [

                [

                    novoId,

                    nomeVoluntario
                ]
            ]
        }
    });


    // ==========================================
    // 3. LOCALIZAR LINHA
    // ==========================================

    const voluntariosTurma =
        await sheets.spreadsheets.values.get({

            spreadsheetId:
            spreadsheetId,

            range:
                `${nomeAbaTurma}!A2:B`
        });


    const linhasVoluntarios =
        voluntariosTurma.data.values || [];


    let linhaNovoVoluntario = -1;


    for (
        let i = 0;
        i < linhasVoluntarios.length;
        i++
    ) {

        if (
            Number(
                linhasVoluntarios[i][0]
            ) ===
            Number(novoId)
        ) {

            linhaNovoVoluntario =
                i + 2;

            break;
        }
    }


    // ==========================================
    // 4. DEIXAR CHAMADAS BRANCAS
    // ==========================================

    if (
        linhaNovoVoluntario !== -1
    ) {

        await limparLinhaChamada(

            abaTurma.properties.sheetId,

            nomeAbaTurma,

            linhaNovoVoluntario
        );
    }


    // ==========================================
    // 5. LOG
    // ==========================================

    await registrarAlteracao(

        administrador.idAdministrador,

        administrador.nome,

        administrador.email,

        nomeAbaTurma,

        "ADICIONAR",

        `Voluntário ${nomeVoluntario} (ID ${novoId}) adicionado na turma ${nomeAbaTurma}.`
    );


    return {

        id:
        novoId,

        nome:
        nomeVoluntario,

        turma:
        nomeAbaTurma,

        curso:
        cursoFinal
    };
}


// ==========================================
// LOCALIZAR VOLUNTÁRIO PELO ID
// ==========================================

async function localizarVoluntarioPorId(
    id,
    planilha
) {

    const idVoluntario =
        Number(id);


    for (
        const aba of
    planilha.data.sheets || []
        ) {

        const nomeAba =
            aba.properties.title;


        if (
            !nomeAba.startsWith(
                "VoluntariosTurma"
            )
        ) {

            continue;
        }


        const resposta =
            await sheets.spreadsheets.values.get({

                spreadsheetId:
                spreadsheetId,

                range:
                    `${nomeAba}!A2:B`
            });


        const linhas =
            resposta.data.values || [];


        for (
            let i = 0;
            i < linhas.length;
            i++
        ) {

            const idAtual =
                Number(linhas[i][0]);


            if (
                idAtual ===
                idVoluntario
            ) {

                return {

                    id:
                    idAtual,

                    nome:
                        linhas[i][1],

                    turma:
                    nomeAba,

                    linha:
                        i + 2,

                    sheetId:
                    aba.properties.sheetId
                };
            }
        }
    }


    return null;
}


// ==========================================
// ATUALIZAR VOLUNTÁRIO
// ==========================================
// Compatível com a rota:
// atualizarVoluntario(
//     id,
//     novoNome,
//     novaTurma,
//     curso,
//     outroCurso,
//     administrador
// )
// ==========================================

async function atualizarVoluntario(
    id,
    novoNome,
    novaTurma,
    curso,
    outroCurso,
    administrador
) {

    if (
        id === undefined ||
        id === null ||
        String(id).trim() === ""
    ) {

        throw new Error(
            "ID do voluntário é obrigatório."
        );
    }


    const idVoluntario =
        Number(id);


    if (
        !Number.isInteger(idVoluntario) ||
        idVoluntario <= 0
    ) {

        throw new Error(
            "ID do voluntário inválido."
        );
    }


    const planilha =
        await sheets.spreadsheets.get({

            spreadsheetId:
            spreadsheetId
        });


    const voluntarioEncontrado =
        await localizarVoluntarioPorId(
            idVoluntario,
            planilha
        );


    if (
        !voluntarioEncontrado
    ) {

        throw new Error(
            "Voluntário não encontrado."
        );
    }


    let alterouAlgumaCoisa = false;


    // ==========================================
    // ATUALIZAR NOME
    // ==========================================

    if (
        novoNome !== undefined &&
        novoNome !== null &&
        String(novoNome).trim() !== ""
    ) {

        const nomeNovo =
            String(novoNome).trim();


        if (
            nomeNovo.toLowerCase() !==
            String(
                voluntarioEncontrado.nome
            )
                .trim()
                .toLowerCase()
        ) {

            await sheets.spreadsheets.values.update({

                spreadsheetId:
                spreadsheetId,

                range:
                    `${voluntarioEncontrado.turma}!B${voluntarioEncontrado.linha}`,

                valueInputOption:
                    "USER_ENTERED",

                requestBody: {

                    values: [
                        [nomeNovo]
                    ]
                }
            });


            // --------------------------------------
            // ATUALIZAR HISTÓRICO
            // --------------------------------------

            const historico =
                await sheets.spreadsheets.values.get({

                    spreadsheetId:
                    spreadsheetId,

                    range:
                        "HistoricoVoluntario!A:F"
                });


            const linhasHistorico =
                historico.data.values || [];


            for (
                let i = 1;
                i < linhasHistorico.length;
                i++
            ) {

                const idHistorico =
                    Number(
                        linhasHistorico[i][0]
                    );


                if (
                    idHistorico ===
                    idVoluntario
                ) {

                    await sheets.spreadsheets.values.update({

                        spreadsheetId:
                        spreadsheetId,

                        range:
                            `HistoricoVoluntario!B${i + 1}`,

                        valueInputOption:
                            "USER_ENTERED",

                        requestBody: {

                            values: [
                                [nomeNovo]
                            ]
                        }
                    });

                    break;
                }
            }


            await registrarAlteracao(

                administrador.idAdministrador,

                administrador.nome,

                administrador.email,

                voluntarioEncontrado.turma,

                "ATUALIZAR",

                `Voluntário ID ${idVoluntario} teve o nome alterado de ${voluntarioEncontrado.nome} para ${nomeNovo}.`
            );


            voluntarioEncontrado.nome =
                nomeNovo;

            alterouAlgumaCoisa = true;
        }
    }


    // ==========================================
    // ATUALIZAR CURSO
    // ==========================================

    if (
        curso !== undefined &&
        curso !== null &&
        String(curso).trim() !== ""
    ) {

        const cursosPermitidos = [

            "Sistemas de Informação",
            "ADS",
            "Fisioterapia",
            "Enfermagem",
            "Medicina",
            "Outro"
        ];


        if (
            !cursosPermitidos.includes(
                curso
            )
        ) {

            throw new Error(
                "Curso inválido."
            );
        }


        if (
            curso === "Outro" &&
            (
                outroCurso === undefined ||
                outroCurso === null ||
                String(outroCurso).trim() === ""
            )
        ) {

            throw new Error(
                "Informe o nome do outro curso."
            );
        }


        let cursoFinal =
            curso;


        if (
            curso === "Outro"
        ) {

            cursoFinal =
                String(outroCurso).trim();
        }


        // --------------------------------------
        // BUSCAR HISTÓRICO
        // --------------------------------------

        const historico =
            await sheets.spreadsheets.values.get({

                spreadsheetId:
                spreadsheetId,

                range:
                    "HistoricoVoluntario!A:F"
            });


        const linhasHistorico =
            historico.data.values || [];


        for (
            let i = 1;
            i < linhasHistorico.length;
            i++
        ) {

            const idHistorico =
                Number(
                    linhasHistorico[i][0]
                );


            if (
                idHistorico ===
                idVoluntario
            ) {

                await sheets.spreadsheets.values.update({

                    spreadsheetId:
                    spreadsheetId,

                    range:
                        `HistoricoVoluntario!C${i + 1}`,

                    valueInputOption:
                        "USER_ENTERED",

                    requestBody: {

                        values: [
                            [cursoFinal]
                        ]
                    }
                });

                break;
            }
        }


        await registrarAlteracao(

            administrador.idAdministrador,

            administrador.nome,

            administrador.email,

            voluntarioEncontrado.turma,

            "ATUALIZAR",

            `Voluntário ID ${idVoluntario} teve o curso alterado para ${cursoFinal}.`
        );


        alterouAlgumaCoisa = true;
    }


    // ==========================================
    // ATUALIZAR TURMA
    // ==========================================

    if (
        novaTurma !== undefined &&
        novaTurma !== null &&
        String(novaTurma).trim() !== ""
    ) {

        const nomeNovaAba =
            obterNomeAbaVoluntarios(
                novaTurma
            );


        if (
            nomeNovaAba !==
            voluntarioEncontrado.turma
        ) {

            const abaNovaTurma =
                planilha.data.sheets.find(

                    sheet =>
                        sheet.properties.title ===
                        nomeNovaAba
                );


            if (!abaNovaTurma) {

                throw new Error(
                    "Turma não encontrada."
                );
            }


            // --------------------------------------
            // 1. ADICIONAR NA NOVA TURMA
            // --------------------------------------

            await sheets.spreadsheets.values.append({

                spreadsheetId:
                spreadsheetId,

                range:
                    `${nomeNovaAba}!A:B`,

                valueInputOption:
                    "USER_ENTERED",

                insertDataOption:
                    "INSERT_ROWS",

                requestBody: {

                    values: [

                        [

                            voluntarioEncontrado.id,

                            voluntarioEncontrado.nome
                        ]
                    ]
                }
            });


            // --------------------------------------
            // 2. LOCALIZAR NOVA LINHA
            // --------------------------------------

            const novosVoluntarios =
                await sheets.spreadsheets.values.get({

                    spreadsheetId:
                    spreadsheetId,

                    range:
                        `${nomeNovaAba}!A2:B`
                });


            const linhasNovos =
                novosVoluntarios.data.values || [];


            let linhaNova = -1;


            for (
                let i = 0;
                i < linhasNovos.length;
                i++
            ) {

                if (
                    Number(
                        linhasNovos[i][0]
                    ) ===
                    Number(
                        voluntarioEncontrado.id
                    )
                ) {

                    linhaNova =
                        i + 2;

                    break;
                }
            }


            // --------------------------------------
            // 3. LIMPAR CHAMADAS DA NOVA TURMA
            // --------------------------------------

            if (
                linhaNova !== -1
            ) {

                await limparLinhaChamada(

                    abaNovaTurma.properties.sheetId,

                    nomeNovaAba,

                    linhaNova
                );
            }


            // --------------------------------------
            // 4. REMOVER DA TURMA ANTIGA
            // --------------------------------------

            const abaAntiga =
                planilha.data.sheets.find(

                    sheet =>
                        sheet.properties.title ===
                        voluntarioEncontrado.turma
                );


            if (!abaAntiga) {

                throw new Error(
                    "Aba antiga não encontrada."
                );
            }


            await sheets.spreadsheets.batchUpdate({

                spreadsheetId:
                spreadsheetId,

                requestBody: {

                    requests: [

                        {

                            deleteDimension: {

                                range: {

                                    sheetId:
                                    abaAntiga.properties.sheetId,

                                    dimension:
                                        "ROWS",

                                    startIndex:
                                        voluntarioEncontrado.linha - 1,

                                    endIndex:
                                    voluntarioEncontrado.linha
                                }
                            }
                        }
                    ]
                }
            });


            // --------------------------------------
            // 5. ZERAR HISTÓRICO
            // --------------------------------------

            const historico =
                await sheets.spreadsheets.values.get({

                    spreadsheetId:
                    spreadsheetId,

                    range:
                        "HistoricoVoluntario!A:F"
                });


            const linhasHistorico =
                historico.data.values || [];


            for (
                let i = 1;
                i < linhasHistorico.length;
                i++
            ) {

                const idHistorico =
                    Number(
                        linhasHistorico[i][0]
                    );


                if (
                    idHistorico ===
                    voluntarioEncontrado.id
                ) {

                    const cursoAtual =
                        linhasHistorico[i][2] || "";


                    await sheets.spreadsheets.values.update({

                        spreadsheetId:
                        spreadsheetId,

                        range:
                            `HistoricoVoluntario!B${i + 1}:F${i + 1}`,

                        valueInputOption:
                            "USER_ENTERED",

                        requestBody: {

                            values: [

                                [

                                    voluntarioEncontrado.nome,

                                    cursoAtual,

                                    0,

                                    0,

                                    nomeNovaAba
                                ]
                            ]
                        }
                    });

                    break;
                }
            }


            await registrarAlteracao(

                administrador.idAdministrador,

                administrador.nome,

                administrador.email,

                nomeNovaAba,

                "ATUALIZAR",

                `Voluntário ID ${voluntarioEncontrado.id} teve a turma alterada de ${voluntarioEncontrado.turma} para ${nomeNovaAba}. Histórico de presença zerado.`
            );


            voluntarioEncontrado.turma =
                nomeNovaAba;

            alterouAlgumaCoisa = true;
        }
    }


    if (
        !alterouAlgumaCoisa
    ) {

        throw new Error(
            "Nenhuma alteração foi informada."
        );
    }


    return {

        id:
        voluntarioEncontrado.id,

        nome:
        voluntarioEncontrado.nome,

        turma:
        voluntarioEncontrado.turma
    };
}


async function listarVoluntarios() {
    const planilha = await sheets.spreadsheets.get({
        spreadsheetId
    });

    const abas = planilha.data.sheets || [];

    const historico = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "HistoricoVoluntario!A:F"
    });

    const linhasHistorico = historico.data.values || [];

    const cursosPorId = new Map();

    for (let i = 1; i < linhasHistorico.length; i++) {
        const linha = linhasHistorico[i];

        if (linha.length < 6) {
            continue;
        }

        const id = Number(linha[0]);

        if (!Number.isInteger(id)) {
            continue;
        }

        cursosPorId.set(id, {
            curso: String(linha[2] || ""),
            turma: String(linha[5] || "")
        });
    }

    const voluntarios = [];
    const idsAdicionados = new Set();

    for (const aba of abas) {
        const nomeAba = aba.properties.title;

        if (!nomeAba.startsWith("VoluntariosTurma")) {
            continue;
        }

        const resposta = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `${nomeAba}!A2:ZZ`
        });

        const linhas = resposta.data.values || [];

        for (const linha of linhas) {
            if (linha.length < 2) {
                continue;
            }

            const id = Number(linha[0]);
            const nome = String(linha[1] || "").trim();

            if (!Number.isInteger(id) || id <= 0 || !nome) {
                continue;
            }

            if (idsAdicionados.has(id)) {
                continue;
            }

            idsAdicionados.add(id);

            let presencas = 0;
            let faltas = 0;

            for (let coluna = 2; coluna < linha.length; coluna++) {
                const status = String(linha[coluna] || "")
                    .trim()
                    .toUpperCase();

                if (status === "P") {
                    presencas++;
                } else if (status === "A") {
                    faltas++;
                }
            }

            const dadosHistorico = cursosPorId.get(id);

            voluntarios.push({
                id: id,
                nome: nome,
                turma: dadosHistorico
                    ? dadosHistorico.turma
                    : nomeAba,
                curso: dadosHistorico
                    ? dadosHistorico.curso
                    : "",
                presencas: presencas,
                faltas: faltas,
                status: dadosHistorico
                    ? dadosHistorico.turma
                    : nomeAba
            });
        }
    }

    voluntarios.sort((a, b) =>
        a.nome.localeCompare(b.nome, "pt-BR")
    );

    return voluntarios;
}
// ==========================================
// PESQUISAR VOLUNTÁRIO
// ==========================================

async function pesquisarVoluntario(
    valor
) {

    if (
        valor === undefined ||
        valor === null ||
        String(valor).trim() === ""
    ) {

        throw new Error(
            "Informe o nome ou ID do voluntário."
        );
    }


    const valorPesquisa =
        String(valor).trim();


    const planilha =
        await sheets.spreadsheets.get({

            spreadsheetId:
            spreadsheetId
        });


    const voluntariosEncontrados = [];


    // ==========================================
    // PESQUISAR NAS TURMAS
    // ==========================================

    for (
        const aba of
    planilha.data.sheets || []
        ) {

        const turma =
            aba.properties.title;


        if (
            !turma.startsWith(
                "VoluntariosTurma"
            )
        ) {

            continue;
        }


        // ALTERADO:
        // Busca também as colunas das datas.

        const resposta =
            await sheets.spreadsheets.values.get({

                spreadsheetId:
                spreadsheetId,

                range:
                    `${turma}!A2:ZZ`
            });


        const linhas =
            resposta.data.values || [];


        for (
            const linha of linhas
            ) {

            const id =
                Number(linha[0]);


            const nome =
                String(
                    linha[1] || ""
                );


            const idCorresponde =
                String(id) ===
                valorPesquisa;


            const nomeCorresponde =
                nome
                    .toLowerCase()
                    .includes(
                        valorPesquisa.toLowerCase()
                    );


            if (
                idCorresponde ||
                nomeCorresponde
            ) {

                let presencas = 0;
                let faltas = 0;


                // ==========================================
                // CONTAR AS CHAMADAS
                // C EM DIANTE = DATAS
                // ==========================================

                for (
                    let coluna = 2;
                    coluna < linha.length;
                    coluna++
                ) {

                    const status =
                        String(
                            linha[coluna] || ""
                        )
                            .trim()
                            .toUpperCase();


                    if (
                        status === "P"
                    ) {

                        presencas++;

                    } else if (
                        status === "A"
                    ) {

                        faltas++;
                    }
                }


                voluntariosEncontrados.push({

                    id:
                    id,

                    nome:
                    nome,

                    turma:
                    turma,

                    presencas:
                    presencas,

                    faltas:
                    faltas
                });
            }
        }
    }


    // ==========================================
    // NENHUM RESULTADO
    // ==========================================

    if (
        voluntariosEncontrados.length === 0
    ) {

        throw new Error(
            "Voluntário não encontrado."
        );
    }


    // ==========================================
    // BUSCAR HISTÓRICO
    // ==========================================

    const historico =
        await sheets.spreadsheets.values.get({

            spreadsheetId:
            spreadsheetId,

            range:
                "HistoricoVoluntario!A:F"
        });


    const linhasHistorico =
        historico.data.values || [];


    for (
        const voluntario of
        voluntariosEncontrados
        ) {

        voluntario.curso = "";

        voluntario.status =
            voluntario.turma;


        for (
            let i = 1;
            i < linhasHistorico.length;
            i++
        ) {

            const idHistorico =
                Number(
                    linhasHistorico[i][0]
                );


            if (
                idHistorico ===
                voluntario.id
            ) {

                voluntario.curso =
                    linhasHistorico[i][2] || "";


                voluntario.status =
                    linhasHistorico[i][5] ||
                    voluntario.turma;


                break;
            }
        }
    }


    // ==========================================
    // ORDEM ALFABÉTICA
    // ==========================================

    voluntariosEncontrados.sort(
        (a, b) =>
            String(a.nome)
                .localeCompare(
                    String(b.nome),
                    "pt-BR",
                    {
                        sensitivity: "base"
                    }
                )
    );


    // ==========================================
    // IMPORTANTE:
    // O FRONT ESPERA UM ARRAY
    // ==========================================

    return voluntariosEncontrados;
}


// ==========================================
// EXCLUIR VOLUNTÁRIO
// ==========================================

async function excluirVoluntario(
    id,
    administrador
) {

    if (
        id === undefined ||
        id === null ||
        String(id).trim() === ""
    ) {

        throw new Error(
            "ID do voluntário é obrigatório."
        );
    }


    const idVoluntario =
        Number(id);


    if (
        !Number.isInteger(idVoluntario) ||
        idVoluntario <= 0
    ) {

        throw new Error(
            "ID do voluntário inválido."
        );
    }


    const planilha =
        await sheets.spreadsheets.get({

            spreadsheetId:
            spreadsheetId
        });


    const voluntarioEncontrado =
        await localizarVoluntarioPorId(
            idVoluntario,
            planilha
        );


    if (
        !voluntarioEncontrado
    ) {

        throw new Error(
            "Voluntário não encontrado."
        );
    }


    // ==========================================
    // REMOVE DA TURMA
    // ==========================================

    await sheets.spreadsheets.batchUpdate({

        spreadsheetId:
        spreadsheetId,

        requestBody: {

            requests: [

                {

                    deleteDimension: {

                        range: {

                            sheetId:
                            voluntarioEncontrado.sheetId,

                            dimension:
                                "ROWS",

                            startIndex:
                                voluntarioEncontrado.linha - 1,

                            endIndex:
                            voluntarioEncontrado.linha
                        }
                    }
                }
            ]
        }
    });


    // ==========================================
    // REMOVE DO HISTÓRICO
    // ==========================================

    const historico =
        await sheets.spreadsheets.values.get({

            spreadsheetId:
            spreadsheetId,

            range:
                "HistoricoVoluntario!A:F"
        });


    const linhasHistorico =
        historico.data.values || [];


    const abaHistorico =
        planilha.data.sheets.find(

            sheet =>
                sheet.properties.title ===
                "HistoricoVoluntario"
        );


    if (!abaHistorico) {

        throw new Error(
            "Aba HistoricoVoluntario não encontrada."
        );
    }


    for (
        let i = 1;
        i < linhasHistorico.length;
        i++
    ) {

        const idHistorico =
            Number(
                linhasHistorico[i][0]
            );


        if (
            idHistorico ===
            idVoluntario
        ) {

            await sheets.spreadsheets.batchUpdate({

                spreadsheetId:
                spreadsheetId,

                requestBody: {

                    requests: [

                        {

                            deleteDimension: {

                                range: {

                                    sheetId:
                                    abaHistorico
                                        .properties
                                        .sheetId,

                                    dimension:
                                        "ROWS",

                                    startIndex:
                                    i,

                                    endIndex:
                                        i + 1
                                }
                            }
                        }
                    ]
                }
            });

            break;
        }
    }


    // ==========================================
    // LOG
    // ==========================================

    await registrarAlteracao(

        administrador.idAdministrador,

        administrador.nome,

        administrador.email,

        voluntarioEncontrado.turma,

        "REMOVER",

        `Voluntário ${voluntarioEncontrado.nome} (ID ${idVoluntario}) removido da turma ${voluntarioEncontrado.turma}.`
    );


    return {

        id:
        idVoluntario,

        nome:
        voluntarioEncontrado.nome,

        turma:
        voluntarioEncontrado.turma,

        mensagem:
            "Voluntário excluído com sucesso."
    };
}


// ==========================================
// REGISTRAR CHAMADA DOS VOLUNTÁRIOS
// ==========================================

async function registrarChamadaVoluntarios(
    turma,
    chamada,
    administrador
) {

    if (!turma) {

        throw new Error(
            "Turma é obrigatória."
        );
    }


    if (
        !Array.isArray(chamada) ||
        chamada.length === 0
    ) {

        throw new Error(
            "A chamada deve possuir voluntários."
        );
    }


    const planilha =
        await sheets.spreadsheets.get({

            spreadsheetId:
            spreadsheetId
        });


    const nomeAbaTurma =
        obterNomeAbaVoluntarios(
            turma
        );


    const abaTurma =
        planilha.data.sheets.find(

            sheet =>
                sheet.properties.title ===
                nomeAbaTurma
        );


    if (!abaTurma) {

        throw new Error(
            "Turma não encontrada."
        );
    }


    const respostaVoluntarios =
        await sheets.spreadsheets.values.get({

            spreadsheetId:
            spreadsheetId,

            range:
                `${nomeAbaTurma}!A2:B`
        });


    const voluntarios =
        respostaVoluntarios.data.values || [];


    // ==========================================
    // VALIDAR VOLUNTÁRIOS
    // ==========================================

    for (
        const registro of
        chamada
        ) {

        const existe =
            voluntarios.some(

                linha =>
                    Number(linha[0]) ===
                    Number(registro.id)
            );


        if (!existe) {

            throw new Error(

                `Voluntário com ID ${registro.id} não encontrado na turma.`
            );
        }
    }


    // ==========================================
    // DATA
    // ==========================================

    const agora =
        new Date();


    const data =
        agora.toLocaleDateString(
            "pt-BR"
        );


    // ==========================================
    // CABEÇALHO
    // ==========================================

    const primeiraLinha =
        await sheets.spreadsheets.values.get({

            spreadsheetId:
            spreadsheetId,

            range:
                `${nomeAbaTurma}!1:1`
        });


    const cabecalho =
        primeiraLinha.data.values?.[0] ||
        [];


    if (
        cabecalho.includes(data)
    ) {

        throw new Error(
            "A chamada dessa data já foi registrada."
        );
    }


    // ==========================================
    // NOVA COLUNA
    // ==========================================

    const novaColuna =
        cabecalho.length + 1;


    const letraColuna =
        numeroParaColuna(
            novaColuna
        );


    // ==========================================
    // DATA NO CABEÇALHO
    // ==========================================

    await sheets.spreadsheets.values.update({

        spreadsheetId:
        spreadsheetId,

        range:
            `${nomeAbaTurma}!${letraColuna}1`,

        valueInputOption:
            "USER_ENTERED",

        requestBody: {

            values: [
                [data]
            ]
        }
    });


    // ==========================================
    // ORDENAR POR NOME
    // ==========================================

    const chamadaOrdenada =
        [...chamada].sort(
            (a, b) => {

                const nomeA =
                    String(
                        voluntarios.find(
                            linha =>
                                Number(linha[0]) ===
                                Number(a.id)
                        )?.[1] || ""
                    );

                const nomeB =
                    String(
                        voluntarios.find(
                            linha =>
                                Number(linha[0]) ===
                                Number(b.id)
                        )?.[1] || ""
                    );


                return nomeA.localeCompare(
                    nomeB,
                    "pt-BR",
                    {
                        sensitivity:
                            "base"
                    }
                );
            }
        );


    // ==========================================
    // HISTÓRICO
    // ==========================================

    const historico =
        await sheets.spreadsheets.values.get({

            spreadsheetId:
            spreadsheetId,

            range:
                "HistoricoVoluntario!A:F"
        });


    const linhasHistorico =
        historico.data.values || [];


    // ==========================================
    // REGISTRAR CADA VOLUNTÁRIO
    // ==========================================

    for (
        const registro of
        chamadaOrdenada
        ) {

        const linhaVoluntario =
            voluntarios.findIndex(

                linha =>
                    Number(linha[0]) ===
                    Number(registro.id)
            );


        if (
            linhaVoluntario === -1
        ) {

            continue;
        }


        const numeroLinha =
            linhaVoluntario + 2;


        // --------------------------------------
        // PRESENÇA / FALTA
        // --------------------------------------

        const status =
            registro.presente
                ? "P"
                : "A";


        await sheets.spreadsheets.values.update({

            spreadsheetId:
            spreadsheetId,

            range:
                `${nomeAbaTurma}!${letraColuna}${numeroLinha}`,

            valueInputOption:
                "USER_ENTERED",

            requestBody: {

                values: [
                    [status]
                ]
            }
        });


        // --------------------------------------
        // ATUALIZAR HISTÓRICO
        // --------------------------------------

        for (
            let i = 1;
            i < linhasHistorico.length;
            i++
        ) {

            const idHistorico =
                Number(
                    linhasHistorico[i][0]
                );


            if (
                idHistorico ===
                Number(registro.id)
            ) {

                let faltas =
                    Number(
                        linhasHistorico[i][3]
                    ) || 0;


                let presencas =
                    Number(
                        linhasHistorico[i][4]
                    ) || 0;


                if (
                    registro.presente
                ) {

                    presencas++;

                } else {

                    faltas++;
                }


                await sheets.spreadsheets.values.update({

                    spreadsheetId:
                    spreadsheetId,

                    range:
                        `HistoricoVoluntario!D${i + 1}:E${i + 1}`,

                    valueInputOption:
                        "USER_ENTERED",

                    requestBody: {

                        values: [

                            [
                                faltas,
                                presencas
                            ]
                        ]
                    }
                });


                break;
            }
        }
    }


    // ==========================================
    // CORES
    // ==========================================

    const requests = [];


    for (
        const registro of
        chamadaOrdenada
        ) {

        const linhaVoluntario =
            voluntarios.findIndex(

                linha =>
                    Number(linha[0]) ===
                    Number(registro.id)
            );


        if (
            linhaVoluntario === -1
        ) {

            continue;
        }


        const numeroLinha =
            linhaVoluntario + 2;


        const cor =
            registro.presente

                ? {

                    red: 0.56,

                    green: 0.93,

                    blue: 0.56

                }

                : {

                    red: 0.95,

                    green: 0.55,

                    blue: 0.55

                };


        requests.push({

            repeatCell: {

                range: {

                    sheetId:
                    abaTurma.properties.sheetId,

                    startRowIndex:
                        numeroLinha - 1,

                    endRowIndex:
                    numeroLinha,

                    startColumnIndex:
                        novaColuna - 1,

                    endColumnIndex:
                    novaColuna
                },

                cell: {

                    userEnteredFormat: {

                        backgroundColor:
                        cor,

                        horizontalAlignment:
                            "CENTER",

                        verticalAlignment:
                            "MIDDLE"
                    }
                },

                fields:
                    "userEnteredFormat.backgroundColor," +
                    "userEnteredFormat.horizontalAlignment," +
                    "userEnteredFormat.verticalAlignment"
            }
        });
    }


    if (
        requests.length > 0
    ) {

        await sheets.spreadsheets.batchUpdate({

            spreadsheetId:
            spreadsheetId,

            requestBody: {

                requests:
                requests
            }
        });
    }


    // ==========================================
    // LOG
    // ==========================================

    await registrarAlteracao(

        administrador.idAdministrador,

        administrador.nome,

        administrador.email,

        nomeAbaTurma,

        "ADICIONAR",

        `Chamada dos voluntários da turma ${nomeAbaTurma} registrada em ${data}.`
    );


    // ==========================================
    // RETORNO
    // ==========================================

    return {

        turma:
        nomeAbaTurma,

        data:
        data,

        mensagem:
            "Chamada dos voluntários registrada com sucesso."
    };
}


// ==========================================
// EXPORTAÇÕES
// ==========================================

export {

    adicionarVoluntario,

    atualizarVoluntario,

    listarVoluntarios,

    pesquisarVoluntario,

    excluirVoluntario,

    registrarChamadaVoluntarios

};