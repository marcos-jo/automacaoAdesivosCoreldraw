import fs from "node:fs";
import path from "node:path";

import { parseScript } from "./src/parser.js";
import { generateScript } from "./src/generator.js";


const caminho = path.resolve(
    "./modelos/Adesivo 33x48_Modelo.js"
);


console.log("Lendo:", caminho);


const originalScript = fs.readFileSync(
    caminho,
    "utf8"
);


const model = parseScript(originalScript);


console.log("Script carregado:");
console.log(model.nome);


console.log("\nVariáveis encontradas:");

for (const section of model.sections) {

    console.log("\n[" + section.nome + "]");

    for (const control of section.controls) {

        console.log(
            control.variable,
            "=",
            control.value
        );

    }

}

console.log("\nAlterando tamanhoHorizontal para 35...");

const controle = model.sections
    .flatMap(section => section.controls)
    .find(control =>
        control.variable === "tamanhoHorizontal"
    );


controle.value = 35;


const novoScript = generateScript(
    originalScript,
    model
);


console.log("\nResultado:");

console.log(novoScript);

function alterar(model, variable, value) {

    const control = model.sections
        .flatMap(section => section.controls)
        .find(control =>
            control.variable === variable
        );

    if (!control) {

        throw new Error(
            `Controle não encontrado: ${variable}`
        );

    }

    control.value = value;

}