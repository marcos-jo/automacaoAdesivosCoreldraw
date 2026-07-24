import { parseScript } from "./parser.js";
import fs from "node:fs";
import path from "node:path";

const MODELOS_DIR = path.resolve("./modelos");

export function listarScripts() {

    const arquivos = fs.readdirSync(MODELOS_DIR);

    return arquivos.filter(arquivo =>
        arquivo.toLowerCase().endsWith(".js")
    );

}

export function carregarScript(nomeArquivo) {

    const caminho = path.join(MODELOS_DIR, nomeArquivo);

    if (!fs.existsSync(caminho)) {

        throw new Error("Script não encontrado: " + nomeArquivo);

    }

    const conteudo = fs.readFileSync(caminho, "utf8");

    const model = parseScript(conteudo);

    model.nome = nomeArquivo;
    model.caminho = caminho;

    return model;

}