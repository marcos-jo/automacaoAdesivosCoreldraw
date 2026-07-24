import { ScriptModel, Section, Control } from "./model.js";
export function parseScript(script) {

    const linhas = script.split(/\r?\n/);

    const model = new ScriptModel();

    let dentroConfig = false;
    let sectionAtual = null;
    let metadata = {};

    for (let numeroLinha = 0; numeroLinha < linhas.length; numeroLinha++) {

        const linha = linhas[numeroLinha].trim();

        console.log(linha);

        if (linha === "//<CONFIG>") {
            dentroConfig = true;
            continue;
        }

        if (linha === "//</CONFIG>") {
            break;
        }

        if (!dentroConfig)
            continue;

        if (linha === "")
            continue;

        //---------------------------------
        // SECTION
        //---------------------------------

        if (linha.startsWith("//@section")) {

            const nome = linha.substring(10).trim();

            sectionAtual = new Section(nome);

            model.addSection(sectionAtual);

            continue;
        }

        //---------------------------------
        // METADATA
        //---------------------------------

        if (linha.startsWith("//@")) {

            const texto = linha.substring(3).trim();

            const indiceEspaco = texto.indexOf(" ");

            if (indiceEspaco === -1) {

                metadata[texto] = true;

            } else {

                const chave = texto.substring(0, indiceEspaco);

                const valor = texto.substring(indiceEspaco + 1).trim();

                metadata[chave] = valor;

            }

            continue;
        }

        //---------------------------------
        // VAR
        //---------------------------------

        if (linha.startsWith("var ")) {

            if (!sectionAtual) {

                throw new Error(
                    `Linha ${numeroLinha + 1}: variável encontrada antes de qualquer //@section`
                );

            }

            const controle = criarControle(linha, metadata);

            sectionAtual.addControl(controle);
            
            metadata = {};

        }

    }

    return model;

}

function criarControle(linha, metadata) {

    //---------------------------------
    // remove "var"
    //---------------------------------

    let texto = linha.substring(3).trim();

    //---------------------------------
    // separa nome e valor
    //---------------------------------

    const indiceIgual = texto.indexOf("=");

    if (indiceIgual === -1) {

        throw new Error("Variável inválida: " + linha);

    }

    const nome = texto.substring(0, indiceIgual).trim();

    let valor = texto.substring(indiceIgual + 1).trim();

    //---------------------------------
    // remove comentários
    //---------------------------------

    const indiceComentario = valor.indexOf("//");

    if (indiceComentario !== -1) {

        valor = valor.substring(0, indiceComentario).trim();

    }

    //---------------------------------
    // remove ;
    //---------------------------------

    if (valor.endsWith(";")) {

        valor = valor.substring(0, valor.length - 1).trim();

    }

    //---------------------------------
    // mm(...)
    //---------------------------------

    let usaMM = false;

    if (valor.startsWith("mm(") && valor.endsWith(")")) {

        usaMM = true;

        valor = valor.substring(3, valor.length - 1).trim();

    }

    //---------------------------------
    // boolean
    //---------------------------------

    if (valor === "true") {

        valor = true;

    } else if (valor === "false") {

        valor = false;

    }

    //---------------------------------
    // number
    //---------------------------------

    else if (!isNaN(Number(valor))) {

        valor = Number(valor);

    }

    //---------------------------------
    // string
    //---------------------------------

    else {

        valor = valor.replace(/^"/, "").replace(/"$/, "");

    }

    const control = new Control();

    control.initialize(nome, valor, usaMM);
    control.applyMetadata(metadata);

    return control;

}