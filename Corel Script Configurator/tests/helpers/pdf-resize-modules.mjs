import { readFile } from "node:fs/promises";
const url = caminho => new URL(caminho, import.meta.url).href;
const carregar = async (caminho, trocas) => {
    let codigo = await readFile(new URL(caminho, import.meta.url), "utf8");
    for (const [de, para] of Object.entries(trocas)) codigo = codigo.replaceAll(JSON.stringify(de), JSON.stringify(para));
    return "data:text/javascript;base64," + Buffer.from(codigo).toString("base64");
};
const biblioteca = url("../../public/vendor/pdf-lib/pdf-lib.js");
export const motorUrl = await carregar("../../public/js/tools/pdf-resize-engine.js", {
    "/vendor/pdf-lib/pdf-lib.js": biblioteca, "./pdf-formats.js": url("../../public/js/tools/pdf-formats.js")
});
export const workspaceUrl = await carregar("../../public/js/tools/pdf-workspace.js", {
    "/vendor/pdf-lib/pdf-lib.js": biblioteca, "./pdf-resize-engine.js": motorUrl
});
export const { PDFWorkspace } = await import(workspaceUrl);
export const { calcularTransformacao, mmParaPontos } = await import(motorUrl);
