export function renderQRCode() {

    const app = document.getElementById("app");

    app.innerHTML = "";

    const titulo = document.createElement("h1");

    titulo.textContent = "Gerador de QR Codes";

    app.appendChild(titulo);

    const texto = document.createElement("p");

    texto.textContent =
        "Importe uma planilha para gerar QR Codes de contatos.";

    app.appendChild(texto);

}