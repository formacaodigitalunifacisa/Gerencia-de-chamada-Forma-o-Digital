import express from "express";

import {
    adicionarTurma,
    listarTurmas
} from "../services/turmaService.js";

import autenticarAdministrador from "../middleware/autenticarAdministrador.js";


const router = express.Router();


// ==========================================
// ADICIONAR TURMA
// ==========================================

router.post("/", autenticarAdministrador, async (req, res) => {

    try {

        console.log("POST de turma recebido!");
        console.log("Dados:", req.body);

        const { turma } = req.body;


        const novaTurma = await adicionarTurma(
            turma,
            req.administrador
        );


        console.log("Turma cadastrada:", novaTurma);

        res.status(201).json(novaTurma);

    } catch (erro) {

        console.error("Erro ao cadastrar turma:", erro);

        res.status(400).json({
            mensagem: erro.message
        });
    }

});


// ==========================================
// LISTAR TURMAS
// ==========================================

router.get("/", autenticarAdministrador, async (req, res) => {

    try {

        const turmas = await listarTurmas();

        res.json(turmas);

    } catch (erro) {

        console.error("Erro ao listar turmas:", erro);

        res.status(500).json({
            mensagem: erro.message
        });

    }

});


export default router;