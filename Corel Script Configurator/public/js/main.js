import { InterfacePrincipal } from "./app/shell.js";
import { AbasAplicacao } from "./app/tabs.js";
import { MenusAplicacao } from "./app/menus.js";
import { JanelaExportacao } from "./app/export-dialog.js";
import { renderQRCode } from "./tools/qrcodes.js";
import { renderPDFs } from "./tools/pdfs.js";

let exportacao;
const modulos = [
    { id: "pdfs", titulo: "PDFs", descricao: "Organize, edite e prepare documentos para impressão.", montar(painel) {
        const pdfs = renderPDFs(painel, { exportacao });
        return { destruir: pdfs, suspender: pdfs.suspender, retomar: pdfs.retomar,
            confirmarFechamento: pdfs.confirmarSaida, podeAlternar: pdfs.podeAlternar, menus: pdfs.menus };
    } },
    { id: "qrcodes", titulo: "QR Codes", descricao: "Módulo de contatos e QR Codes em desenvolvimento.", montar(painel) { renderQRCode(painel); return {}; } }
];
const interfacePrincipal = new InterfacePrincipal(document.getElementById("app-shell"), modulos);
exportacao = new JanelaExportacao(interfacePrincipal.raiz);
const abas = new AbasAplicacao(modulos, interfacePrincipal);
// A fila já informa o erro; comandos antigos/indisponíveis preservam a aba atual.
interfacePrincipal.aoAtivar = id => abas.ativar(id).catch(() => false);
interfacePrincipal.aoFechar = id => abas.fechar(id);
const menus = new MenusAplicacao(document.getElementById("app-menus"), texto => interfacePrincipal.informar(texto));
interfacePrincipal.aoExibir = aba => menus.renderizar(aba, { fechar:() => abas.fechar(aba.id), inicio:() => abas.ativar("inicio"),
    verificarAtualizacoes: globalThis.window?.desktop4PrintShops?.verificarAtualizacoes });
interfacePrincipal.aoExibir(abas.abas.get("inicio"));
