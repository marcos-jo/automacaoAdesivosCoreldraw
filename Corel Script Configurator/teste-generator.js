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

console.log("\nAlterando...");

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

alterar(model, "tamanhoHorizontal", 35);
alterar(model, "tamanhoVertical", 48);
alterar(model, "quantidadeDeCopiasHorizontal", 3);
alterar(model, "quantidadeDeCopiasVertical", 4);
alterar(model, "corteRedondo", false);
alterar(model, "cortePersonalizado", true);
alterar(model, "posicaoInicialX", 50);
alterar(model, "posicaoInicialY", 37);

                    console.log("\nVALORES ANTES DO GENERATOR:");

                    for (const section of model.sections) {

                        for (const control of section.controls) {

                            console.log(
                                control.variable,
                                "=",
                                control.value
                            );

                        }

                    }


const novoScript = generateScript(
    originalScript,
    model
);

console.log("\nRESULTADO GERADO:\n");

console.log(novoScript);