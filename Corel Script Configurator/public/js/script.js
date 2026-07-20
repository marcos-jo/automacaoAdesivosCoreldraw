import { carregarModelo } from "./api.js";
import { renderizarFormulario } from "./renderer.js";

async function iniciar(){

    const modelo = await carregarModelo();

    renderizarFormulario(modelo);

}

iniciar();