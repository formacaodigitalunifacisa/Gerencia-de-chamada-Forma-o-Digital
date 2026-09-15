import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import session from "express-session";

import alunoRoutes from "./routes/alunoRoutes.js";
import voluntarioRoutes from "./routes/voluntarioRoutes.js";
import loginRoutes from "./routes/loginRoutes.js";
import administradorRoutes from "./routes/administradorRoutes.js";
import turmaRoutes from "./routes/turmaRoutes.js";

dotenv.config();

const app = express();

const PORT = process.env.PORT || 3000;


// ==========================================
// CONFIGURAÇÕES
// ==========================================

app.use(cors({
    origin: "gerencia-formacao-digital-front.vercel.app",
    credentials: true
}));

app.use(express.json());


// ==========================================
// SESSÃO DO ADMINISTRADOR
// ==========================================

app.use(session({
    secret: "formacao-digital-segredo",
    resave: false,
    saveUninitialized: false,
    cookie: {
        secure: false
    }
}));


// ==========================================
// ROTAS
// ==========================================

app.use("/api/alunos", alunoRoutes);

app.use("/api/voluntarios", voluntarioRoutes);

app.use("/api/login", loginRoutes);

app.use("/api/administrador", administradorRoutes);

app.use("/api/turmas", turmaRoutes);


// ==========================================
// ROTA PRINCIPAL
// ==========================================

app.get("/", (req, res) => {

    res.json({
        mensagem: "Backend Formação Digital funcionando!"
    });

});


// ==========================================
// INICIAR SERVIDOR
// ==========================================

app.listen(PORT, () => {

    console.log(
        `Servidor da Formação Digital rodando na porta ${PORT}`
    );

});