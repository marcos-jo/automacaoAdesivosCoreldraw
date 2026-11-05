import { instalarAtalhos } from "./app/shortcuts.js";
import { iniciarTema } from "./app/theme.js";
import { InterfacePrincipal } from "./app/shell.js";
import { AbasAplicacao } from "./app/tabs.js";
import { MenusAplicacao } from "./app/menus.js";
import { JanelaExportacao } from "./app/export-dialog.js";
import { renderQRCode } from "./tools/qrcodes.js";
import { renderPDFs } from "./tools/pdfs.js";
import { instalarAtualizacoes } from "./app/desktop-updates.js";
import { instalarAtualizacaoAssistida } from "./app/desktop-assisted.js";
import { instalarAberturaWindows } from "./app/desktop-open.js";
import { JanelaConfiguracaoImpressao } from "./app/desktop-print-ui.js";

iniciarTema();
let exportacao;
let montagemImpressao = null;
const modulos = [
    { id: "pdfs", titulo: "PDFs", descricao: "Organize, edite e prepare documentos para impressão.", montar(painel) {
        const pdfs = renderPDFs(painel, { exportacao, imprimirMontagem: globalThis.window?.desktop4PrintShops?.imprimirDocumento ? async result => {
            montagemImpressao=result; configuracaoImpressao.get('scale').value='100'; await configuracaoImpressao.open();
        } : null });
        return { destruir: pdfs, suspender: pdfs.suspender, retomar: pdfs.retomar,
            confirmarFechamento: pdfs.confirmarSaida, podeAlternar: pdfs.podeAlternar, menus: pdfs.menus, estadoTrabalho: pdfs.estadoTrabalho,
            abrirArquivo:pdfs.abrirArquivo, adicionarArquivos:pdfs.adicionarArquivos,temDocumento:pdfs.temDocumento,
            abrirLeitorExterno:pdfs.abrirLeitorExterno,enviarTesteRetido:pdfs.enviarTesteRetido,imprimirDocumento:pdfs.imprimirDocumento,fonteImpressao:pdfs.fonteImpressao,prepararPreviaImpressao:pdfs.prepararPreviaImpressao };
    } },
    { id: "qrcodes", titulo: "QR Codes", descricao: "Módulo de contatos e QR Codes em desenvolvimento.", montar(painel) { renderQRCode(painel); return {}; } }
];
const interfacePrincipal = new InterfacePrincipal(document.getElementById("app-shell"), modulos);
exportacao = new JanelaExportacao(interfacePrincipal.raiz);
const abas = new AbasAplicacao(modulos, interfacePrincipal);
const configuracaoImpressao=globalThis.window?.desktop4PrintShops?.configuracaoImpressao
    ? new JanelaConfiguracaoImpressao(interfacePrincipal.raiz,window.desktop4PrintShops,{
        imprimir:(state,scale)=>montagemImpressao ? window.desktop4PrintShops.imprimirDocumento({session:state.session,revision:state.revision,scale,name:'Montagem 4PrintShops',operation:crypto.randomUUID().replaceAll('-','')},montagemImpressao.bytes) : abas.abas.get('pdfs')?.controle.imprimirDocumento(state,scale),
        fonte:()=>montagemImpressao ? {width:montagemImpressao.plan.folha.largura,height:montagemImpressao.plan.folha.altura,page:1,pages:montagemImpressao.plan.totalFolhas,name:'Montagem 4PrintShops'} : abas.abas.get('pdfs')?.controle.fonteImpressao(),
        prepararPrevia:()=>montagemImpressao ? {bytes:montagemImpressao.bytes.slice(),page:1} : abas.abas.get('pdfs')?.controle.prepararPreviaImpressao(),
        enviarTeste:(state,scale)=>abas.abas.get('pdfs')?.controle.enviarTesteRetido(state,scale),
        abrirSintetico:async()=>{
            const bytes=await window.desktop4PrintShops.ensaioRetido({kind:'fixture'});
            return abas.abas.get('pdfs')?.controle.abrirArquivo(new File([bytes],'4PrintShops Sintético Hold.pdf',{type:'application/pdf'}));
        }
    }) : null;
configuracaoImpressao?.dialog.addEventListener('close',()=>{montagemImpressao=null;});
// A fila já informa o erro; comandos antigos/indisponíveis preservam a aba atual.
interfacePrincipal.aoAtivar = id => abas.ativar(id).catch(() => false);
interfacePrincipal.aoAbrirPDF = async () => {
    if(await abas.ativar('pdfs'))interfacePrincipal.conteudo.querySelector('#pdf-arquivo')?.click();
};
interfacePrincipal.aoFechar = id => abas.fechar(id);
const menus = new MenusAplicacao(document.getElementById("app-menus"), texto => interfacePrincipal.informar(texto));
interfacePrincipal.aoExibir = aba => menus.renderizar(aba, { fechar:() => abas.fechar(aba.id), inicio:() => abas.ativar("inicio"),
    abrirPDF:()=>interfacePrincipal.aoAbrirPDF(),
    verificarAtualizacoes: globalThis.window?.desktop4PrintShops?.verificarAtualizacoes,
    configurarImpressao:configuracaoImpressao?()=>configuracaoImpressao.open():null,
    abrirLeitorExterno:globalThis.window?.desktop4PrintShops?.abrirCopiaExterna&&aba.controle?.abrirLeitorExterno?()=>aba.controle.abrirLeitorExterno():null,
    verificarAssistida: globalThis.window?.desktop4PrintShops?.atualizacaoAssistida ? () => window.desktop4PrintShops.atualizacaoAssistida('check') : null });
interfacePrincipal.aoExibir(abas.abas.get("inicio"));
instalarAtualizacoes(interfacePrincipal.raiz, () => [...abas.abas.values()]);
instalarAtualizacaoAssistida(interfacePrincipal.raiz);
instalarAberturaWindows(abas,interfacePrincipal);

instalarAtalhos({
    available(action){
        const pdf=abas.abas.get('pdfs')?.controle;
        if([...abas.abas.values()].some(tab=>tab.controle.estadoTrabalho?.().busy))return false;
        return action==='open'||!!(pdf?.temDocumento()&&document.querySelector('#workspace-pdfs')&&!document.querySelector('#workspace-pdfs').hidden);
    },
    async execute(action){
        if(action==='open')return interfacePrincipal.aoAbrirPDF();
        if(action==='print'){if(configuracaoImpressao)await configuracaoImpressao.open();else interfacePrincipal.informar('Impressão direta disponível na versão Desktop.');return;}
        interfacePrincipal.conteudo.querySelector('#pdf-salvar-copia')?.click();
    }
});