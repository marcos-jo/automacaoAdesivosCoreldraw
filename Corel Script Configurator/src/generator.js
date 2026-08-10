export function generateScript(originalScript, model) {

    let resultado = originalScript;

    for (const section of model.sections) {

        for (const control of section.controls) {

            resultado = substituirVariavel(
                resultado,
                control
            );

        }

    }

    return resultado;

}


function substituirVariavel(script, control) {

    const nome = control.variable;

    let novoValor;

    //---------------------------------
    // Formata o valor
    //---------------------------------

    if (control.usaMM) {

        novoValor = `mm(${control.value})`;

    }

    else if (typeof control.value === "boolean") {

        novoValor = control.value
            ? "true"
            : "false";

    }

    else if (typeof control.value === "number") {

        novoValor = String(control.value);

    }

    else {

        novoValor = `"${control.value}"`;

    }

    //---------------------------------
    // Substitui a variável
    //---------------------------------

    const regex = new RegExp(
        `(var\\s+${nome}\\s*=\\s*)([^;]+)(;)`
    );

    return script.replace(
        regex,
        `$1${novoValor}$3`
    );

}