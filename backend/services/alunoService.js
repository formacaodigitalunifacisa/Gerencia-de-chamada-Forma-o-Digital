import { sheets, spreadsheetId } from "../config/googleSheets.js";
import { registrarAlteracao } from "./logService.js";


// ==========================================
// CONVERTER NÚMERO PARA LETRA DA COLUNA
// ==========================================

function numeroParaColuna(numero) {

    let coluna = "";

    while (numero > 0) {

        const resto = (numero - 1) % 26;

        coluna =
            String.fromCharCode(65 + resto) +
            coluna;

        numero =
            Math.floor((numero - 1) / 26);
    }

    return coluna;
}


// ==========================================
// LIMPAR E DEIXAR BRANCA A LINHA DE CHAMADA
// ==========================================

async function limparLinhaChamada(
    sheetId,
    nomeAba,
    numeroLinha
) {

    const primeiraLinha =
        await sheets.spreadsheets.values.get({
            spreadsheetId: spreadsheetId,
            range: `${nomeAba}!1:1`
        });

    const cabecalho =
        primeiraLinha.data.values?.[0] || [];


    // Se só existem ID e Nome,
    // ainda não existem chamadas.

    if (cabecalho.length <= 2) {
        return;
    }


    const ultimaColuna =
        numeroParaColuna(cabecalho.length);


    // ------------------------------------------
    // LIMPA OS VALORES DAS CHAMADAS
    // ------------------------------------------

    await sheets.spreadsheets.values.clear({
        spreadsheetId: spreadsheetId,
        range:
            `${nomeAba}!C${numeroLinha}:${ultimaColuna}${numeroLinha}`
    });


    // ------------------------------------------
    // DEIXA AS CÉLULAS BRANCAS
    // ------------------------------------------

    await sheets.spreadsheets.batchUpdate({

        spreadsheetId: spreadsheetId,

        requestBody: {

            requests: [

                {
                    repeatCell: {

                        range: {

                            sheetId: sheetId,

                            startRowIndex:
                                numeroLinha - 1,

                            endRowIndex:
                            numeroLinha,

                            startColumnIndex: 2,

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
// BUSCAR PRÓXIMO ID DOS ALUNOS
// ==========================================
//
// O ID nunca será reutilizado.
//
// Exemplo:
//
// ID 1
// ID 2
// ID 3
// ID 4
// ID 5 -> excluído
// ID 6
//
// Próximo ID = 7
//
// O ID excluído deixa de existir nas turmas
// e no histórico, mas continua registrado
// no LogAlteracoes.
//
// Por isso procuramos o maior ID em:
//
// 1. Histórico
// 2. Turmas
// 3. LogAlteracoes
//
// O próximo ID será sempre maior que todos.
//

async function obterProximoIdAluno(planilha) {

    let maiorId = 0;


    // ==========================================
    // 1. PROCURA NOS HISTÓRICOS
    // ==========================================

    try {

        const historico =
            await sheets.spreadsheets.values.get({

                spreadsheetId: spreadsheetId,

                range:
                    "HistoricoAluno!A2:A"
            });


        const linhasHistorico =
            historico.data.values || [];


        for (const linha of linhasHistorico) {

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
            "Não foi possível consultar o histórico dos alunos:",
            erro.message
        );
    }


    // ==========================================
    // 2. PROCURA NAS TURMAS
    // ==========================================

    for (const aba of planilha.data.sheets || []) {

        const nomeAba =
            aba.properties.title;


        if (
            !nomeAba.startsWith("Alunos")
        ) {

            continue;
        }


        try {

            const resposta =
                await sheets.spreadsheets.values.get({

                    spreadsheetId: spreadsheetId,

                    range:
                        `${nomeAba}!A2:A`
                });


            const linhas =
                resposta.data.values || [];


            for (const linha of linhas) {

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
    // 3. PROCURA NOS LOGS
    // ==========================================
    //
    // Não pegamos qualquer número do log.
    //
    // Existem números de:
    //
    // - administrador
    // - turma
    // - data
    // - hora
    //
    // Pegamos somente números associados
    // a "Aluno" e "ID".
    //
    // Exemplos reconhecidos:
    //
    // Aluno João (ID 5)
    // Aluno ID 5
    // Aluno Maria - ID 7
    //
    // Também aceita "Aluno" sem depender
    // da acentuação de outras palavras.
    //

    try {

        const logs =
            await sheets.spreadsheets.values.get({

                spreadsheetId: spreadsheetId,

                range:
                    "LogAlteracoes!A2:I"
            });


        const linhasLogs =
            logs.data.values || [];


        for (const linha of linhasLogs) {

            const descricao =
                String(linha[8] || "");


            const correspondencia =
                descricao.match(
                    /\bAluno\b.*?\bID\s*[:#-]?\s*(\d+)/i
                );


            if (correspondencia) {

                const id =
                    Number(correspondencia[1]);


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
            "Não foi possível consultar o LogAlteracoes para o ID dos alunos:",
            erro.message
        );
    }


    return maiorId + 1;
}


// ==========================================
// ADICIONAR ALUNO
// ==========================================

async function adicionarAluno(
    nome,
    turma,
    administrador
) {

    if (!nome || !turma) {

        throw new Error(
            "Nome e turma são obrigatórios."
        );
    }


    const nomeAluno =
        String(nome).trim();

    const nomeTurma =
        String(turma).trim();


    if (!nomeAluno) {

        throw new Error(
            "Nome do aluno é obrigatório."
        );
    }


    if (!nomeTurma) {

        throw new Error(
            "Turma é obrigatória."
        );
    }


    // ==========================================
    // BUSCA AS ABAS
    // ==========================================

    const planilha =
        await sheets.spreadsheets.get({

            spreadsheetId:
            spreadsheetId
        });


    // ==========================================
    // VERIFICA TURMA
    // ==========================================

    const abaTurma =
        planilha.data.sheets.find(

            sheet =>
                sheet.properties.title ===
                nomeTurma
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
        await obterProximoIdAluno(planilha);


    // ==========================================
    // 1. HISTÓRICO PRIMEIRO
    // ==========================================

    await sheets.spreadsheets.values.append({

        spreadsheetId:
        spreadsheetId,

        range:
            "HistoricoAluno!A:E",

        valueInputOption:
            "USER_ENTERED",

        insertDataOption:
            "INSERT_ROWS",

        requestBody: {

            values: [

                [
                    novoId,
                    nomeAluno,
                    0,
                    0,
                    nomeTurma
                ]
            ]
        }
    });


    // ==========================================
    // 2. ADICIONA NA TURMA
    // ==========================================

    await sheets.spreadsheets.values.append({

        spreadsheetId:
        spreadsheetId,

        range:
            `${nomeTurma}!A:B`,

        valueInputOption:
            "USER_ENTERED",

        insertDataOption:
            "INSERT_ROWS",

        requestBody: {

            values: [

                [
                    novoId,
                    nomeAluno
                ]
            ]
        }
    });


    // ==========================================
    // 3. LOCALIZA A LINHA DO NOVO ALUNO
    // ==========================================

    const alunosTurma =
        await sheets.spreadsheets.values.get({

            spreadsheetId:
            spreadsheetId,

            range:
                `${nomeTurma}!A2:B`
        });


    const linhasAlunos =
        alunosTurma.data.values || [];


    let linhaNovoAluno = -1;


    for (
        let i = 0;
        i < linhasAlunos.length;
        i++
    ) {

        if (
            Number(linhasAlunos[i][0]) ===
            Number(novoId)
        ) {

            linhaNovoAluno =
                i + 2;

            break;
        }
    }


    // ==========================================
    // 4. GARANTE CÉLULAS DE CHAMADA BRANCAS
    // ==========================================

    if (
        linhaNovoAluno !== -1
    ) {

        await limparLinhaChamada(

            abaTurma.properties.sheetId,

            nomeTurma,

            linhaNovoAluno
        );
    }


    // ==========================================
    // 5. LOG
    // ==========================================

    await registrarAlteracao(

        administrador.idAdministrador,

        administrador.nome,

        administrador.email,

        nomeTurma,

        "ADICIONAR",

        `Aluno ${nomeAluno} (ID ${novoId}) adicionado na turma ${nomeTurma}.`
    );


    // ==========================================
    // RETORNO
    // ==========================================

    return {

        id:
        novoId,

        nome:
        nomeAluno,

        turma:
        nomeTurma
    };
}


// ==========================================
// ATUALIZAR ALUNO
// ==========================================

async function atualizarAluno(
    nomeAtual,
    opcao,
    novoNome,
    novaTurma,
    administrador
) {

    if (
        !nomeAtual ||
        !opcao
    ) {

        throw new Error(
            "Nome atual e opção de atualização são obrigatórios."
        );
    }


    const nomeAtualLimpo =
        String(nomeAtual).trim();


    const opcaoLimpa =
        String(opcao).trim().toLowerCase();


    // ==========================================
    // BUSCA AS ABAS
    // ==========================================

    const planilha =
        await sheets.spreadsheets.get({

            spreadsheetId:
            spreadsheetId
        });


    // ==========================================
    // PROCURA ALUNO
    // ==========================================

    let alunoEncontrado = null;


    for (const aba of planilha.data.sheets || []) {

        const turma =
            aba.properties.title;


        if (
            !turma.startsWith("Alunos")
        ) {

            continue;
        }


        const resposta =
            await sheets.spreadsheets.values.get({

                spreadsheetId:
                spreadsheetId,

                range:
                    `${turma}!A2:B`
            });


        const linhas =
            resposta.data.values || [];


        for (
            let i = 0;
            i < linhas.length;
            i++
        ) {

            const id =
                Number(linhas[i][0]);

            const nome =
                linhas[i][1];


            if (
                String(nome).trim().toLowerCase() ===
                nomeAtualLimpo.toLowerCase()
            ) {

                alunoEncontrado = {

                    id:
                    id,

                    nome:
                    nome,

                    turma:
                    turma,

                    linha:
                        i + 2,

                    sheetId:
                    aba.properties.sheetId
                };

                break;
            }
        }


        if (alunoEncontrado) {
            break;
        }
    }


    if (!alunoEncontrado) {

        throw new Error(
            "Aluno não encontrado."
        );
    }


    // ==========================================
    // ATUALIZAR NOME
    // ==========================================

    if (
        opcaoLimpa === "nome"
    ) {

        if (!novoNome) {

            throw new Error(
                "Informe o novo nome."
            );
        }


        const nomeNovo =
            String(novoNome).trim();


        if (!nomeNovo) {

            throw new Error(
                "O novo nome não pode estar vazio."
            );
        }


        // --------------------------------------
        // ATUALIZA TURMA
        // --------------------------------------

        await sheets.spreadsheets.values.update({

            spreadsheetId:
            spreadsheetId,

            range:
                `${alunoEncontrado.turma}!B${alunoEncontrado.linha}`,

            valueInputOption:
                "USER_ENTERED",

            requestBody: {

                values: [
                    [nomeNovo]
                ]
            }
        });


        // --------------------------------------
        // ATUALIZA HISTÓRICO
        // --------------------------------------

        const historico =
            await sheets.spreadsheets.values.get({

                spreadsheetId:
                spreadsheetId,

                range:
                    "HistoricoAluno!A:E"
            });


        const linhasHistorico =
            historico.data.values || [];


        for (
            let i = 1;
            i < linhasHistorico.length;
            i++
        ) {

            const idHistorico =
                Number(linhasHistorico[i][0]);


            if (
                idHistorico ===
                alunoEncontrado.id
            ) {

                await sheets.spreadsheets.values.update({

                    spreadsheetId:
                    spreadsheetId,

                    range:
                        `HistoricoAluno!B${i + 1}`,

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


        // --------------------------------------
        // LOG
        // --------------------------------------

        await registrarAlteracao(

            administrador.idAdministrador,

            administrador.nome,

            administrador.email,

            alunoEncontrado.turma,

            "ATUALIZAR",

            `Aluno ID ${alunoEncontrado.id} teve o nome alterado de ${alunoEncontrado.nome} para ${nomeNovo}.`
        );


        return {

            id:
            alunoEncontrado.id,

            nome:
            nomeNovo,

            turma:
            alunoEncontrado.turma
        };
    }


    // ==========================================
    // ATUALIZAR TURMA
    // ==========================================

    if (
        opcaoLimpa === "turma"
    ) {

        if (!novaTurma) {

            throw new Error(
                "Informe a nova turma."
            );
        }


        const novaTurmaLimpa =
            String(novaTurma).trim();


        // --------------------------------------
        // VERIFICA NOVA TURMA
        // --------------------------------------

        const abaNovaTurma =
            planilha.data.sheets.find(

                sheet =>
                    sheet.properties.title ===
                    novaTurmaLimpa
            );


        if (!abaNovaTurma) {

            throw new Error(
                "Turma não encontrada."
            );
        }


        // --------------------------------------
        // MESMA TURMA
        // --------------------------------------

        if (
            alunoEncontrado.turma ===
            novaTurmaLimpa
        ) {

            throw new Error(
                "O aluno já está nessa turma."
            );
        }


        // --------------------------------------
        // 1. ADICIONA NA NOVA TURMA
        // --------------------------------------

        await sheets.spreadsheets.values.append({

            spreadsheetId:
            spreadsheetId,

            range:
                `${novaTurmaLimpa}!A:B`,

            valueInputOption:
                "USER_ENTERED",

            insertDataOption:
                "INSERT_ROWS",

            requestBody: {

                values: [

                    [
                        alunoEncontrado.id,
                        alunoEncontrado.nome
                    ]
                ]
            }
        });


        // --------------------------------------
        // 2. BUSCA LINHA NOVA
        // --------------------------------------

        const alunosNovaTurma =
            await sheets.spreadsheets.values.get({

                spreadsheetId:
                spreadsheetId,

                range:
                    `${novaTurmaLimpa}!A2:B`
            });


        const linhasNovaTurma =
            alunosNovaTurma.data.values || [];


        let linhaNovoAluno = -1;


        for (
            let i = 0;
            i < linhasNovaTurma.length;
            i++
        ) {

            if (
                Number(linhasNovaTurma[i][0]) ===
                Number(alunoEncontrado.id)
            ) {

                linhaNovoAluno =
                    i + 2;

                break;
            }
        }


        // --------------------------------------
        // 3. LIMPA CÉLULAS DA CHAMADA
        // --------------------------------------
        //
        // O aluno começa na nova turma
        // sem presença ou falta da turma anterior.
        //
        // As células de chamadas ficam brancas.
        //

        if (
            linhaNovoAluno !== -1
        ) {

            await limparLinhaChamada(

                abaNovaTurma.properties.sheetId,

                novaTurmaLimpa,

                linhaNovoAluno
            );
        }


        // --------------------------------------
        // 4. REMOVE DA TURMA ANTIGA
        // --------------------------------------

        const abaAntiga =
            planilha.data.sheets.find(

                sheet =>
                    sheet.properties.title ===
                    alunoEncontrado.turma
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
                                    alunoEncontrado.linha - 1,

                                endIndex:
                                alunoEncontrado.linha
                            }
                        }
                    }
                ]
            }
        });


        // --------------------------------------
        // 5. ATUALIZA HISTÓRICO
        // --------------------------------------

        const historico =
            await sheets.spreadsheets.values.get({

                spreadsheetId:
                spreadsheetId,

                range:
                    "HistoricoAluno!A:E"
            });


        const linhasHistorico =
            historico.data.values || [];


        for (
            let i = 1;
            i < linhasHistorico.length;
            i++
        ) {

            const idHistorico =
                Number(linhasHistorico[i][0]);


            if (
                idHistorico ===
                alunoEncontrado.id
            ) {

                await sheets.spreadsheets.values.update({

                    spreadsheetId:
                    spreadsheetId,

                    range:
                        `HistoricoAluno!B${i + 1}:E${i + 1}`,

                    valueInputOption:
                        "USER_ENTERED",

                    requestBody: {

                        values: [

                            [
                                alunoEncontrado.nome,
                                0,
                                0,
                                novaTurmaLimpa
                            ]
                        ]
                    }
                });

                break;
            }
        }


        // --------------------------------------
        // 6. LOG
        // --------------------------------------

        await registrarAlteracao(

            administrador.idAdministrador,

            administrador.nome,

            administrador.email,

            novaTurmaLimpa,

            "ATUALIZAR",

            `Aluno ID ${alunoEncontrado.id} teve a turma alterada de ${alunoEncontrado.turma} para ${novaTurmaLimpa}. Histórico de presença zerado.`
        );


        return {

            id:
            alunoEncontrado.id,

            nome:
            alunoEncontrado.nome,

            turma:
            novaTurmaLimpa
        };
    }


    // ==========================================
    // OPÇÃO INVÁLIDA
    // ==========================================

    throw new Error(
        "Opção de atualização inválida. Escolha nome ou turma."
    );
}


/// ==========================================
// LISTAR ALUNOS
// ==========================================

async function listarAlunos() {

    // ==========================================
    // BUSCA TODOS OS ALUNOS DO HISTÓRICO
    // ==========================================

    const historico =
        await sheets.spreadsheets.values.get({

            spreadsheetId:
            spreadsheetId,

            range:
                "HistoricoAluno!A:E"
        });


    const linhasHistorico =
        historico.data.values || [];


    // ==========================================
    // BUSCA AS ABAS DA PLANILHA
    // ==========================================

    const planilha =
        await sheets.spreadsheets.get({

            spreadsheetId:
            spreadsheetId
        });


    const alunos = [];


    // ==========================================
    // EVITA ALUNOS DUPLICADOS
    // ==========================================

    const idsAdicionados =
        new Set();


    // ==========================================
    // GUARDA OS DADOS DAS TURMAS
    // ==========================================

    const dadosTurmas =
        new Map();


    // ==========================================
    // PERCORRE TODOS OS ALUNOS
    // ==========================================

    for (
        let i = 1;
        i < linhasHistorico.length;
        i++
    ) {

        const linha =
            linhasHistorico[i];


        if (
            linha.length < 5
        ) {

            continue;
        }


        const id =
            Number(linha[0]);


        const nome =
            String(linha[1] || "");


        const turma =
            String(linha[4] || "");


        if (
            !Number.isInteger(id)
        ) {

            continue;
        }


        // ==========================================
        // NÃO ADICIONA O MESMO ID DUAS VEZES
        // ==========================================

        if (
            idsAdicionados.has(id)
        ) {

            continue;
        }


        idsAdicionados.add(id);


        // ==========================================
        // PRESENÇAS E FALTAS
        // ==========================================

        let presencas = 0;

        let faltas = 0;


        // ==========================================
        // BUSCA OS DADOS DA TURMA
        // ==========================================

        if (
            turma &&
            turma.startsWith("Alunos")
        ) {

            if (
                !dadosTurmas.has(turma)
            ) {

                try {

                    const respostaTurma =
                        await sheets.spreadsheets.values.get({

                            spreadsheetId:
                            spreadsheetId,

                            range:
                                `${turma}!A2:ZZ`
                        });


                    dadosTurmas.set(

                        turma,

                        respostaTurma.data.values || []
                    );

                } catch (erro) {

                    console.warn(

                        `Não foi possível consultar a turma ${turma}:`,

                        erro.message
                    );


                    dadosTurmas.set(
                        turma,
                        []
                    );
                }
            }


            const linhasTurma =
                dadosTurmas.get(turma) || [];


            // ==========================================
            // PROCURA O ALUNO PELO ID
            // ==========================================

            const linhaAluno =
                linhasTurma.find(

                    linha =>
                        Number(linha[0]) ===
                        id
                );


            if (
                linhaAluno
            ) {

                // ==========================================
                // CONTA P E A DAS DATAS
                // ==========================================

                for (
                    let coluna = 2;
                    coluna < linhaAluno.length;
                    coluna++
                ) {

                    const status =
                        String(
                            linhaAluno[coluna] || ""
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
            }
        }


        // ==========================================
        // ADICIONA ALUNO
        // ==========================================

        alunos.push({

            id:
            id,

            nome:
            nome,

            turma:
            turma,

            presencas:
            presencas,

            faltas:
            faltas,

            status:
            turma
        });
    }


    // ==========================================
    // RETORNO
    // ==========================================

    return alunos;
}



// ==========================================
// PESQUISAR ALUNO
// ==========================================

async function pesquisarAluno(valor) {

    if (
        valor === undefined ||
        valor === null ||
        String(valor).trim() === ""
    ) {

        throw new Error(
            "Informe o nome ou ID do aluno."
        );
    }


    const valorPesquisa =
        String(valor).trim();


    const planilha =
        await sheets.spreadsheets.get({

            spreadsheetId:
            spreadsheetId
        });


    const alunosEncontrados = [];


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
            !turma.startsWith("Alunos")
        ) {

            continue;
        }


        // Busca também as colunas das chamadas
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
                // CONTAR CHAMADAS
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


                alunosEncontrados.push({

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
        alunosEncontrados.length === 0
    ) {

        throw new Error(
            "Aluno não encontrado."
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
                "HistoricoAluno!A:E"
        });


    const linhasHistorico =
        historico.data.values || [];


    // ==========================================
    // COMPLETAR STATUS
    // ==========================================

    for (
        const aluno of
        alunosEncontrados
        ) {

        aluno.status =
            aluno.turma;


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
                aluno.id
            ) {

                aluno.status =
                    linhasHistorico[i][4] ||
                    aluno.turma;

                break;
            }
        }
    }


    // ==========================================
    // ORDEM ALFABÉTICA
    // ==========================================

    alunosEncontrados.sort(

        (a, b) =>
            String(a.nome)
                .localeCompare(
                    String(b.nome),
                    "pt-BR",
                    {
                        sensitivity:
                            "base"
                    }
                )
    );


    // ==========================================
    // RETORNO
    // ==========================================

    return alunosEncontrados;
}


// ==========================================
// EXCLUIR ALUNO
// ==========================================

async function excluirAluno(
    id,
    administrador
) {

    if (
        id === undefined ||
        id === null ||
        String(id).trim() === ""
    ) {

        throw new Error(
            "ID do aluno é obrigatório."
        );
    }


    const idAluno =
        Number(id);


    if (
        !Number.isInteger(idAluno)
    ) {

        throw new Error(
            "ID do aluno inválido."
        );
    }


    // ==========================================
    // BUSCA AS ABAS
    // ==========================================

    const planilha =
        await sheets.spreadsheets.get({

            spreadsheetId:
            spreadsheetId
        });


    let alunoEncontrado = null;


    // ==========================================
    // PROCURA O ALUNO NAS TURMAS
    // ==========================================

    for (
        const aba of
    planilha.data.sheets || []
        ) {

        const turma =
            aba.properties.title;


        if (
            !turma.startsWith("AlunosTurma")
        ) {

            continue;
        }


        const resposta =
            await sheets.spreadsheets.values.get({

                spreadsheetId:
                spreadsheetId,

                range:
                    `${turma}!A2:B`
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
                idAtual === idAluno
            ) {

                alunoEncontrado = {

                    turma:
                    turma,

                    nome:
                        linhas[i][1],

                    sheetId:
                    aba.properties.sheetId
                };

                break;
            }
        }


        if (
            alunoEncontrado
        ) {

            break;
        }
    }


    // ==========================================
    // ALUNO NÃO ENCONTRADO
    // ==========================================

    if (
        !alunoEncontrado
    ) {

        throw new Error(
            "Aluno não encontrado."
        );
    }


    // ==========================================
    // REMOVE TODAS AS OCORRÊNCIAS DA TURMA
    // ==========================================

    const respostaTurma =
        await sheets.spreadsheets.values.get({

            spreadsheetId:
            spreadsheetId,

            range:
                `${alunoEncontrado.turma}!A2:B`
        });


    const linhasTurma =
        respostaTurma.data.values || [];


    const linhasParaExcluirTurma = [];


    for (
        let i = 0;
        i < linhasTurma.length;
        i++
    ) {

        const idAtual =
            Number(linhasTurma[i][0]);


        if (
            idAtual === idAluno
        ) {

            linhasParaExcluirTurma.push(
                i + 2
            );
        }
    }


    // ==========================================
    // EXCLUI AS LINHAS DA TURMA
    // ==========================================
    //
    // Excluímos de baixo para cima para
    // não alterar os números das linhas.
    //

    for (
        let i =
            linhasParaExcluirTurma.length - 1;
        i >= 0;
        i--
    ) {

        const numeroLinha =
            linhasParaExcluirTurma[i];


        await sheets.spreadsheets.batchUpdate({

            spreadsheetId:
            spreadsheetId,

            requestBody: {

                requests: [

                    {
                        deleteDimension: {

                            range: {

                                sheetId:
                                alunoEncontrado.sheetId,

                                dimension:
                                    "ROWS",

                                startIndex:
                                    numeroLinha - 1,

                                endIndex:
                                numeroLinha
                            }
                        }
                    }
                ]
            }
        });
    }


    // ==========================================
    // BUSCA HISTÓRICO
    // ==========================================

    const historico =
        await sheets.spreadsheets.values.get({

            spreadsheetId:
            spreadsheetId,

            range:
                "HistoricoAluno!A:E"
        });


    const linhasHistorico =
        historico.data.values || [];


    const abaHistorico =
        planilha.data.sheets.find(

            sheet =>
                sheet.properties.title ===
                "HistoricoAluno"
        );


    if (!abaHistorico) {

        throw new Error(
            "Aba HistoricoAluno não encontrada."
        );
    }


    // ==========================================
    // LOCALIZA TODAS AS OCORRÊNCIAS
    // ==========================================

    const linhasParaExcluirHistorico = [];


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
            idHistorico === idAluno
        ) {

            // +1 porque o array começa no índice 0
            // e a planilha começa na linha 1.
            linhasParaExcluirHistorico.push(
                i + 1
            );
        }
    }


    // ==========================================
    // EXCLUI TODAS AS OCORRÊNCIAS DO HISTÓRICO
    // ==========================================

    for (
        let i =
            linhasParaExcluirHistorico.length - 1;
        i >= 0;
        i--
    ) {

        const numeroLinha =
            linhasParaExcluirHistorico[i];


        await sheets.spreadsheets.batchUpdate({

            spreadsheetId:
            spreadsheetId,

            requestBody: {

                requests: [

                    {
                        deleteDimension: {

                            range: {

                                sheetId:
                                abaHistorico.properties.sheetId,

                                dimension:
                                    "ROWS",

                                startIndex:
                                    numeroLinha - 1,

                                endIndex:
                                numeroLinha
                            }
                        }
                    }
                ]
            }
        });
    }


    // ==========================================
    // LOG
    // ==========================================
    //
    // O log NÃO é excluído.
    //
    // Isso garante que o ID nunca seja reutilizado.
    //

    await registrarAlteracao(

        administrador.idAdministrador,

        administrador.nome,

        administrador.email,

        alunoEncontrado.turma,

        "REMOVER",

        `Aluno ${alunoEncontrado.nome} (ID ${idAluno}) removido da turma ${alunoEncontrado.turma}.`
    );


    // ==========================================
    // RETORNO
    // ==========================================

    return {

        id:
        idAluno,

        nome:
        alunoEncontrado.nome,

        turma:
        alunoEncontrado.turma,

        mensagem:
            "Aluno excluído com sucesso."
    };
}


// ==========================================
// REGISTRAR CHAMADA
// ==========================================

async function registrarChamada(
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
            "A chamada deve possuir alunos."
        );
    }


    const nomeTurma =
        String(turma).trim();


    const planilha =
        await sheets.spreadsheets.get({

            spreadsheetId:
            spreadsheetId
        });


    const abaTurma =
        planilha.data.sheets.find(

            sheet =>
                sheet.properties.title ===
                nomeTurma
        );


    if (!abaTurma) {

        throw new Error(
            "Turma não encontrada."
        );
    }


    const respostaAlunos =
        await sheets.spreadsheets.values.get({

            spreadsheetId:
            spreadsheetId,

            range:
                `${nomeTurma}!A2:B`
        });


    const alunos =
        respostaAlunos.data.values || [];


    // ==========================================
    // VALIDA ALUNOS
    // ==========================================

    for (const registro of chamada) {

        const alunoExiste =
            alunos.some(

                linha =>
                    Number(linha[0]) ===
                    Number(registro.id)
            );


        if (!alunoExiste) {

            throw new Error(

                `Aluno com ID ${registro.id} não encontrado na turma.`
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
    // ==========================================ac
    const primeiraLinha =
        await sheets.spreadsheets.values.get({

            spreadsheetId:
            spreadsheetId,

            range:
                `${nomeTurma}!1:1`
        });


    const cabecalho =
        primeiraLinha.data.values?.[0] || [];


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
    // ADICIONA DATA
    // ==========================================

    await sheets.spreadsheets.values.update({

        spreadsheetId:
        spreadsheetId,

        range:
            `${nomeTurma}!${letraColuna}1`,

        valueInputOption:
            "USER_ENTERED",

        requestBody: {

            values: [
                [data]
            ]
        }
    });


    // ==========================================
    // ORDENA CHAMADA
    // ==========================================

    const chamadaOrdenada =
        [...chamada].sort(

            (a, b) =>
                Number(a.id) -
                Number(b.id)
        );


    // ==========================================
    // HISTÓRICO
    // ==========================================

    const historico =
        await sheets.spreadsheets.values.get({

            spreadsheetId:
            spreadsheetId,

            range:
                "HistoricoAluno!A:E"
        });


    const linhasHistorico =
        historico.data.values || [];


    // ==========================================
    // REGISTRA CADA ALUNO
    // ==========================================

    for (
        const registro of chamadaOrdenada
        ) {

        const linhaAluno =
            alunos.findIndex(

                linha =>
                    Number(linha[0]) ===
                    Number(registro.id)
            );


        if (
            linhaAluno === -1
        ) {

            continue;
        }


        const numeroLinha =
            linhaAluno + 2;


        // --------------------------------------
        // P / A
        // --------------------------------------

        await sheets.spreadsheets.values.update({

            spreadsheetId:
            spreadsheetId,

            range:
                `${nomeTurma}!${letraColuna}${numeroLinha}`,

            valueInputOption:
                "USER_ENTERED",

            requestBody: {

                values: [

                    [
                        registro.presente
                            ? "P"
                            : "A"
                    ]
                ]
            }
        });


        // --------------------------------------
        // HISTÓRICO
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

                let presencas =
                    Number(
                        linhasHistorico[i][2]
                    ) || 0;


                let faltas =
                    Number(
                        linhasHistorico[i][3]
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
                        `HistoricoAluno!C${i + 1}:D${i + 1}`,

                    valueInputOption:
                        "USER_ENTERED",

                    requestBody: {

                        values: [

                            [
                                presencas,
                                faltas
                            ]
                        ]
                    }
                });


                break;
            }
        }
    }


    // ==========================================
    // CORES DA CHAMADA
    // ==========================================

    const requests = [];


    for (
        const registro of chamadaOrdenada
        ) {

        const linhaAluno =
            alunos.findIndex(

                linha =>
                    Number(linha[0]) ===
                    Number(registro.id)
            );


        if (
            linhaAluno === -1
        ) {

            continue;
        }


        const numeroLinha =
            linhaAluno + 2;


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

        nomeTurma,

        "CHAMADA",

        `Chamada dos alunos da turma ${nomeTurma} registrada em ${data}.`
    );


    // ==========================================
    // RETORNO
    // ==========================================

    return {

        turma:
        nomeTurma,

        data:
        data,

        mensagem:
            "Chamada registrada com sucesso."
    };
}


// ==========================================
// EXPORTAÇÕES
// ==========================================

export {

    adicionarAluno,

    atualizarAluno,

    listarAlunos,

    pesquisarAluno,

    excluirAluno,

    registrarChamada

};