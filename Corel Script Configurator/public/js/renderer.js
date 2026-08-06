export function renderScript(model) {

    const app = document.getElementById("app");

    app.innerHTML = "";

    for (const section of model.sections) {

        renderSection(app, section);

    }

    updateVisibility();

}

function renderSection(container, section) {

    const fieldset = document.createElement("fieldset");
    fieldset.className = "section";
    fieldset.dataset.section = section.nome;

    const legend = document.createElement("legend");
    legend.className = "section-title";
    legend.textContent = section.nome;

    fieldset.appendChild(legend);

    //---------------------------------
    // Conteúdo da seção
    //---------------------------------

    const content = document.createElement("div");
    content.className = "section-content";

    for (const control of section.controls) {

        renderControl(content, control);

    }

    fieldset.appendChild(content);

    container.appendChild(fieldset);

}

function renderControl(container, control) {

    switch (control.type) {

        case "radio":
            renderRadio(container, control);
            break;

        case "number":
            renderNumber(container, control);
            break;

        default:

            console.warn(
                "Tipo não suportado:",
                control.type
            );

    }

}

function renderRadio(container, control) {

    const wrapper = document.createElement("div");

    wrapper.className = "control radio-control";

if (control.showif) {

    wrapper.dataset.showif = control.showif;

}

    const label = document.createElement("label");
    label.className = "radio-label";

    const input = document.createElement("input");

    input.dataset.variable = control.variable;

    input.addEventListener("change", () => {

        bindControl(control, input);

        updateVisibility();

    });

    input.type = "radio";
    input.name = control.group;
    input.checked = control.value;
    input.value = "true";

    input.className = "radio-input";

    label.appendChild(input);

    label.append(" " + control.label);

    wrapper.appendChild(label);

    if (control.help) {

        const help = document.createElement("small");

        help.className = "help";

        help.textContent = control.help;

        wrapper.appendChild(help);

    }

    container.appendChild(wrapper);

}

function renderNumber(container, control) {

    const wrapper = document.createElement("div");

    if (control.showif) {

        wrapper.dataset.showif = control.showif;

    }
    wrapper.className = "control number-control";

    const label = document.createElement("label");
    label.className = "control-label";
    label.textContent = control.label;

    wrapper.appendChild(label);

    const input = document.createElement("input");

    input.dataset.variable = control.variable;

    input.type = "number";
    input.className = "number-input";

    input.value = control.value;

    input.addEventListener("input", () => {

        atualizarControle(control, input);

    });

    wrapper.appendChild(input);

    if (control.unit) {

        const unit = document.createElement("span");

        unit.className = "unit";

        unit.textContent = control.unit;

        wrapper.appendChild(unit);

    }

    if (control.help) {

        const help = document.createElement("small");

        help.className = "help";

        help.textContent = control.help;

        wrapper.appendChild(help);

    }

    container.appendChild(wrapper);

}


//---------------------------------
// Atualiza o modelo sempre que o
// usuário altera um controle
//---------------------------------

function bindControl(control, input) {

    switch (control.type) {

        case "number":

            control.value = Number(input.value);

            break;

        case "radio":

            control.value = input.checked;

            break;

        default:

            control.value = input.value;

    }

}

function updateVisibility() {

    //---------------------------------
    // Controles com showif
    //---------------------------------

    const controles = document.querySelectorAll("[data-showif]");

    for (const controle of controles) {

        const variavel = controle.dataset.showif;

        const input = document.querySelector(
            `input[data-variable="${variavel}"]`
        );

        if (!input)
            continue;


        let visivel = false;


        if (input.type === "radio") {

            visivel = input.checked;

        }


        if (input.type === "checkbox") {

            visivel = input.checked;

        }


        controle.hidden = !visivel;

    }


    //---------------------------------
    // Oculta sections vazias
    //---------------------------------

    const sections = document.querySelectorAll(".section");


    for (const section of sections) {

        const controlesVisiveis =
            section.querySelectorAll(
                ".control:not([hidden])"
            );


        section.hidden =
            controlesVisiveis.length === 0;

    }

}