export function renderListaScripts(scripts) {

    const app = document.getElementById("app");

    app.innerHTML = "";

    //---------------------------------
    // Título
    //---------------------------------

    const titulo = document.createElement("h2");
    titulo.textContent = "Scripts disponíveis";

    app.appendChild(titulo);

    //---------------------------------
    // Select
    //---------------------------------

    const select = document.createElement("select");

    select.id = "listaScripts";

    for (const script of scripts) {

        const option = document.createElement("option");

        option.value = script;
        option.textContent = script;

        select.appendChild(option);

    }

    app.appendChild(select);

    //---------------------------------
    // Botão Abrir
    //---------------------------------

    const botao = document.createElement("button");

    botao.textContent = "Abrir";

    botao.addEventListener("click", () => {

        window.dispatchEvent(

            new CustomEvent("abrir-script", {

                detail: select.value

            })

        );

    });

    app.appendChild(botao);

}

export function renderScript(model) {

    const app = document.getElementById("app");

    app.innerHTML = "";

    console.log(model);
    console.log(model.sections);

    for (const section of model.sections) {

        const titulo = document.createElement("h2");

        titulo.textContent = section.nome;

        app.appendChild(titulo);

        const hr = document.createElement("hr");

        app.appendChild(hr);

    }

}