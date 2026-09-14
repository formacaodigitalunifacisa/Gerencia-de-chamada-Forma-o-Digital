import express from "express";

import {
    realizarLogin
} from "../services/loginService.js";


const router = express.Router();


// ==========================================
// LOGIN
// ==========================================

router.post("/", async (req, res) => {

    try {

        const { email, senha } = req.body;

        const administrador = await realizarLogin(
            email,
            senha
        );
        req.session.administrador = {
            idAdministrador: administrador.idAdministrador,
            nome: administrador.nome,
            email: administrador.email
        };


        res.json({
            mensagem: "Login realizado com sucesso.",
            administrador: administrador
        });


    } catch (erro) {

        console.error("Erro ao realizar login:", erro);

        res.status(401).json({
            mensagem: erro.message
        });

    }

});


// ==========================================
// LOGOUT
// ==========================================

router.post("/logout", (req, res) => {

    req.session.destroy((erro) => {

        if (erro) {

            console.error("Erro ao encerrar sessão:", erro);

            return res.status(500).json({
                mensagem: "Erro ao sair."
            });
        }

        res.json({
            mensagem: "Logout realizado com sucesso."
        });

    });

});


export default router;