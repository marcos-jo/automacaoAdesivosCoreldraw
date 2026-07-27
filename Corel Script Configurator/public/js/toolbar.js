export function renderToolbar(scripts) {

    const app = document.getElementById("app");

    app.innerHTML = "";

    //---------------------------------
    // Título
    //---------------------------------

    const titulo = document.createElement("h2");
    titulo.textContent = "Scripts disponíveis";

    app.appendChild(titulo);

    const toolbar = document.createElement("div");
    toolbar.className = "toolbar";

    //---------------------------------
    // Select
    //---------------------------------

    const select = document.createElement("select");
    select.id = "listaScripts";

    const placeholder = document.createElement("option");

    placeholder.value = "";
    placeholder.textContent = "Primeiro selecione um modelo...";
    placeholder.selected = true;
    placeholder.disabled = true;

    select.appendChild(placeholder);

    for (const script of scripts) {

        const option = document.createElement("option");

        option.value = script;
        option.textContent = script
            .replace(".js", "")
            .replace("_Modelo", "");

        select.appendChild(option);

    }

    select.addEventListener("change", () => {

    botao.disabled = !select.value;

    });

    //---------------------------------
    // Botão Abrir
    //---------------------------------

    const botao = document.createElement("button");

    botao.textContent = "Abrir";
    botao.disabled = true;

    botao.addEventListener("click", () => {

        window.dispatchEvent(

            new CustomEvent("abrir-script", {

                detail: select.value

            })

        );

    });

    toolbar.appendChild(select);
    toolbar.appendChild(botao);

    // adiciona a toolbar ao app
    app.appendChild(toolbar);

}