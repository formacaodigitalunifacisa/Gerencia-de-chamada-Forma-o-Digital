import express from "express";

import autenticarAdministrador from "../middleware/autenticarAdministrador.js";

const router = express.Router();


// ==========================================
// VERIFICAR ADMINISTRADOR LOGADO
// ==========================================

router.get("/me", autenticarAdministrador, (req, res) => {

    res.json({
        administrador: req.session.administrador
    });

});


export default router;