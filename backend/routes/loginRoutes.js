
import express from "express";
import jwt from "jsonwebtoken";

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


        // ==========================================
        // CRIAR TOKEN
        // ==========================================

        const segredoToken =
            process.env.JWT_SECRET;

        if (!segredoToken) {

            console.error(
                "JWT_SECRET não configurado."
            );

            return res.status(500).json({
                mensagem:
                    "Configuração de autenticação não encontrada."
            });
        }


        const token = jwt.sign(
            {
                idAdministrador:
                    administrador.idAdministrador,

                nome:
                    administrador.nome,

                email:
                    administrador.email
            },
            segredoToken,
            {
                expiresIn: "8h"
            }
        );


        // ==========================================
        // RESPOSTA
        // ==========================================

        res.json({
            mensagem:
                "Login realizado com sucesso.",

            token: token,

            administrador: administrador
        });


    } catch (erro) {

        console.error(
            "Erro ao realizar login:",
            erro
        );

        res.status(401).json({
            mensagem: erro.message
        });

    }

});


// ==========================================
// LOGOUT
// ==========================================

router.post("/logout", (req, res) => {

    res.json({
        mensagem:
            "Logout realizado com sucesso."
    });

});


export default router;
