import express from "express";


import {
    adicionarAluno,
    atualizarAluno,
    listarAlunos,
    pesquisarAluno,
    excluirAluno,
    registrarChamada
} from "../services/alunoService.js";

import autenticarAdministrador from "../middleware/autenticarAdministrador.js";


const router = express.Router();


// ==========================================
// ADICIONAR ALUNO
// ==========================================

router.post("/", autenticarAdministrador, async (req, res) => {

    try {

        console.log("POST de aluno recebido!");
        console.log("Dados:", req.body);

        const { nome, turma } = req.body;

        const aluno = await adicionarAluno(
            nome,
            turma,
            req.administrador
        );

        console.log("Aluno cadastrado:", aluno);

        res.status(201).json(aluno);

    } catch (erro) {

        console.error("Erro ao cadastrar aluno:", erro);

        res.status(400).json({
            mensagem: erro.message
        });
    }
});


// ==========================================
// ATUALIZAR ALUNO
// ==========================================

router.put("/atualizar", autenticarAdministrador, async (req, res) => {

    try {

        const {
            nomeAtual,
            opcao,
            novoNome,
            novaTurma
        } = req.body;


        const aluno = await atualizarAluno(
            nomeAtual,
            opcao,
            novoNome,
            novaTurma,
            req.administrador
        );


        res.json(aluno);

    } catch (erro) {

        console.error("Erro ao atualizar aluno:", erro);

        res.status(400).json({
            mensagem: erro.message
        });
    }
});


// ==========================================
// LISTAR ALUNOS
// ==========================================

router.get("/", async (req, res) => {

    try {

        const alunos = await listarAlunos();

        res.json(alunos);

    } catch (erro) {

        console.error("Erro ao listar alunos:", erro);

        res.status(500).json({
            mensagem: erro.message
        });
    }
});


// ==========================================
// PESQUISAR ALUNO
// ==========================================

router.get("/pesquisar/:valor", async (req, res) => {

    try {

        const { valor } = req.params;

        const alunos = await pesquisarAluno(valor);

        res.json(alunos);

    } catch (erro) {

        console.error("Erro ao pesquisar aluno:", erro);

        res.status(404).json({
            mensagem: erro.message
        });
    }
});


// ==========================================
// FAZER CHAMADA DOS ALUNOS
// ==========================================

router.post("/chamada", autenticarAdministrador, async (req, res) => {

    try {

        const { turma, chamada } = req.body;

        const resultado = await registrarChamada(
            turma,
            chamada,
            req.administrador
        );

        res.json(resultado);

    } catch (erro) {

        console.error("Erro ao registrar chamada:", erro);

        res.status(400).json({
            mensagem: erro.message
        });
    }
});


// ==========================================
// BUSCAR ALUNO POR ID
// ==========================================

router.get("/:id", (req, res) => {

    res.json({
        mensagem: "Rota para buscar aluno"
    });

});


// ==========================================
// EXCLUIR ALUNO
// ==========================================

router.delete("/:id", autenticarAdministrador, async (req, res) => {

    try {

        const { id } = req.params;

        const aluno = await excluirAluno(
            id,
            req.administrador
        );

        res.json(aluno);

    } catch (erro) {

        console.error("Erro ao excluir aluno:", erro);

        res.status(400).json({
            mensagem: erro.message
        });
    }

});


export default router;