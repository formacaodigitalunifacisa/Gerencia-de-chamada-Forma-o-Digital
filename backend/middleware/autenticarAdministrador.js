import jwt from "jsonwebtoken";


// ==========================================
// AUTENTICAR ADMINISTRADOR
// ==========================================

function autenticarAdministrador(req, res, next) {

    try {

        const autorizacao =
            req.headers.authorization;


        // ==========================================
        // VERIFICAR TOKEN
        // ==========================================

        if (!autorizacao) {

            return res.status(401).json({
                mensagem:
                    "Administrador não autenticado."
            });
        }


        // Esperado:
        // Authorization: Bearer TOKEN

        const partes =
            autorizacao.split(" ");


        if (
            partes.length !== 2 ||
            partes[0] !== "Bearer"
        ) {

            return res.status(401).json({
                mensagem:
                    "Token de autenticação inválido."
            });
        }


        const token = partes[1];


        // ==========================================
        // VERIFICAR TOKEN
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


        const administrador =
            jwt.verify(
                token,
                segredoToken
            );


        // ==========================================
        // GUARDAR ADMINISTRADOR NA REQUISIÇÃO
        // ==========================================

        req.administrador =
            administrador;


        next();


    } catch (erro) {

        console.error(
            "Erro ao autenticar administrador:",
            erro.message
        );

        return res.status(401).json({
            mensagem:
                "Token inválido ou expirado."
        });
    }
}


export default autenticarAdministrador;

