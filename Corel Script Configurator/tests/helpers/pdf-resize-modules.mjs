import { readFile } from "node:fs/promises";
const url = caminho => new URL(caminho, import.meta.url).href;
const carregar = async (caminho, trocas) => {
    let codigo = await readFile(new URL(caminho, import.meta.url), "utf8");
    for (const [de, para] of Object.entries(trocas)) codigo = codigo.replaceAll(JSON.stringify(de), JSON.stringify(para));
    return "data:text/javascript;base64," + Buffer.from(codigo).toString("base64");
};
const biblioteca = url("../../public/vendor/pdf-lib/pdf-lib.js");
export const assinaturasUrl = await carregar("../../public/js/tools/pdf-signatures.js", { "/vendor/pdf-lib/pdf-lib.js": biblioteca });
export const copiaImpressaoUrl = await carregar("../../public/js/tools/pdf-print-copy.js", { "/vendor/pdf-lib/pdf-lib.js": biblioteca, "./pdf-signatures.js": assinaturasUrl });
export const impressaoUrl = await carregar("../../public/js/tools/pdf-print-compatibility.js", { "/vendor/pdf-lib/pdf-lib.js": biblioteca });
export const motorUrl = await carregar("../../public/js/tools/pdf-resize-engine.js", {
    "/vendor/pdf-lib/pdf-lib.js": biblioteca, "./pdf-formats.js": url("../../public/js/tools/pdf-formats.js"), "./pdf-print-compatibility.js": impressaoUrl
});
export const workspaceUrl = await carregar("../../public/js/tools/pdf-workspace.js", {
    "/vendor/pdf-lib/pdf-lib.js": biblioteca, "./pdf-resize-engine.js": motorUrl, "./pdf-signatures.js": assinaturasUrl, "./pdf-print-copy.js": copiaImpressaoUrl
});
export const { PDFWorkspace } = await import(workspaceUrl);
export const { calcularTransformacao, mmParaPontos } = await import(motorUrl);
