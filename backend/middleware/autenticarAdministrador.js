function autenticarAdministrador(req, res, next) {

    if (!req.session || !req.session.administrador) {

        return res.status(401).json({
            mensagem: "Administrador não autenticado."
        });
    }

    next();
}


export default autenticarAdministrador;