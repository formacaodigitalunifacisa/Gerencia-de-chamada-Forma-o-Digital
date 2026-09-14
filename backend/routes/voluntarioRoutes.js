import express from "express";

import {
    adicionarVoluntario,
    atualizarVoluntario,
    listarVoluntarios,
    pesquisarVoluntario,
    excluirVoluntario,
    registrarChamadaVoluntarios
} from "../services/voluntarioService.js";

import autenticarAdministrador from "../middleware/autenticarAdministrador.js";

const router = express.Router();


// ==========================================
// ADICIONAR VOLUNTÁRIO
// ==========================================

router.post("/", autenticarAdministrador, async (req, res) => {

    try {

        console.log("POST de voluntário recebido!");
        console.log("Dados:", req.body);

        const {
            nome,
            turma,
            curso,
            outroCurso
        } = req.body;

        const voluntario = await adicionarVoluntario(
            nome,
            turma,
            curso,
            outroCurso,
            req.session.administrador
        );

        console.log("Voluntário cadastrado:", voluntario);

        res.status(201).json(voluntario);

    } catch (erro) {

        console.error("Erro ao cadastrar voluntário:", erro);

        res.status(400).json({
            mensagem: erro.message
        });
    }
});


// ==========================================
// ATUALIZAR VOLUNTÁRIO
// ==========================================

router.put("/atualizar", autenticarAdministrador, async (req, res) => {

    try {

        const {
            id,
            novoNome,
            novaTurma,
            curso,
            outroCurso
        } = req.body;

        const voluntario = await atualizarVoluntario(
            id,
            novoNome,
            novaTurma,
            curso,
            outroCurso,
            req.session.administrador
        );

        res.json(voluntario);

    } catch (erro) {

        console.error("Erro ao atualizar voluntário:", erro);

        res.status(400).json({
            mensagem: erro.message
        });
    }
});


// ==========================================
// LISTAR VOLUNTÁRIOS
// ==========================================

router.get("/", autenticarAdministrador, async (req, res) => {

    try {

        const voluntarios = await listarVoluntarios();

        res.json(voluntarios);

    } catch (erro) {

        console.error("Erro ao listar voluntários:", erro);

        res.status(500).json({
            mensagem: erro.message
        });
    }
});


// ==========================================
// PESQUISAR VOLUNTÁRIO
// ==========================================

router.get("/pesquisar/:valor", autenticarAdministrador, async (req, res) => {

    try {

        const { valor } = req.params;

        const voluntarios = await pesquisarVoluntario(valor);

        res.json(voluntarios);

    } catch (erro) {

        console.error("Erro ao pesquisar voluntário:", erro);

        res.status(404).json({
            mensagem: erro.message
        });
    }
});


// ==========================================
// FAZER CHAMADA DOS VOLUNTÁRIOS
// ==========================================

router.post("/chamada", autenticarAdministrador, async (req, res) => {

    try {

        const {
            turma,
            chamada
        } = req.body;

        const resultado = await registrarChamadaVoluntarios(
            turma,
            chamada,
            req.session.administrador
        );

        res.json(resultado);

    } catch (erro) {

        console.error("Erro ao registrar chamada dos voluntários:", erro);

        res.status(400).json({
            mensagem: erro.message
        });
    }

});


// ==========================================
// EXCLUIR VOLUNTÁRIO
// ==========================================

router.delete("/:id", autenticarAdministrador, async (req, res) => {

    try {

        const { id } = req.params;

        const voluntario = await excluirVoluntario(
            id,
            req.session.administrador
        );

        res.json(voluntario);

    } catch (erro) {

        console.error("Erro ao excluir voluntário:", erro);

        res.status(400).json({
            mensagem: erro.message
        });
    }

});


export default router;