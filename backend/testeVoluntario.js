import { adicionarVoluntario } from "./services/voluntarioService.js";

const administrador = {
    idAdministrador: "1",
    nome: "Administrador Teste",
    email: "admin@teste.com"
};

try {
    const voluntario = await adicionarVoluntario(
        "Carlos M",
        "VoluntariosTurma2",
        "Outro",
        "Engenharia",
        administrador
    );

    console.log("=================================");
    console.log("VOLUNTÁRIO CADASTRADO:");
    console.log(voluntario);
    console.log("=================================");

} catch (erro) {
    console.error("ERRO NO TESTE:");
    console.error(erro);
}