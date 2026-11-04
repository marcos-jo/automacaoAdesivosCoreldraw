import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {PreviaDocumentoImpressao} from '../public/js/app/desktop-print-preview.js';
import {interfaceImpressao} from '../public/js/app/desktop-print-interface.js';
import {JanelaConfiguracaoImpressao} from '../public/js/app/desktop-print-ui.js';
import {assetPath} from '../desktop/security.mjs';

function fixture(){
    const preview=new PreviaDocumentoImpressao({querySelector:()=>({toDataURL:()=>'data:image/png;base64,test'})});
    const calls=[];
    preview.preview={
        documento:{numPages:3},exibicao:{base:{width:841.8898,height:595.2756}},
        preparar:async bytes=>{calls.push(['prepare',bytes]);return {documento:{numPages:3}};},
        adotar:async()=>true,mostrar:async page=>{calls.push(['page',page]);return true;},
        destruir:async()=>calls.push(['destroy'])
    };
    return {preview,calls};
}
test('prévia abre snapshot na página atual e respeita dimensões visuais já giradas',async()=>{
    const {preview,calls}=fixture();await preview.open({bytes:new Uint8Array([1,2]),page:2});
    assert.equal(preview.page,2);assert.equal(preview.source.pages,3);
    assert.ok(Math.abs(preview.source.width-297)<.001);assert.ok(Math.abs(preview.source.height-210)<.001);
    assert.equal(preview.image,'data:image/png;base64,test');assert.deepEqual(calls.map(c=>c[0]),['prepare','page']);
});
test('navegação da prévia não solicita páginas inválidas nem aplica outra rotação',async()=>{
    const {preview,calls}=fixture();await preview.open({bytes:new Uint8Array([1]),page:999});assert.equal(preview.page,3);
    assert.equal(await preview.show(0),false);assert.equal(await preview.show(4),false);
    await preview.show(1);assert.deepEqual(calls.filter(c=>c[0]==='page'),[['page',3],['page',1]]);
});
test('cancelamento durante render não publica imagem antiga e libera o documento',async()=>{
    const {preview,calls}=fixture();let complete;
    preview.preview.mostrar=()=>new Promise(resolve=>{complete=resolve;});
    const pending=preview.show(1);await preview.close();complete(true);
    assert.equal(await pending,false);assert.equal(preview.image,null);assert.equal(preview.source,null);assert.deepEqual(calls,[['destroy']]);
});
test('layout mostra somente opções implementadas e mantém rodapé fora da rolagem',async()=>{
    for(const id of ['printer','properties','copies','pages','scale','orientation','summary','previous','page','next','close','print'])assert.match(interfaceImpressao,new RegExp('id="print-config-'+id+'"'));
    assert.doesNotMatch(interfaceImpressao,/value="fit"|type="radio"|type="number".*copies/);
    assert.match(interfaceImpressao,/Todas as páginas/);assert.match(interfaceImpressao,/Definidas nas propriedades do driver/);
    const footer=interfaceImpressao.slice(interfaceImpressao.indexOf('<footer'));
    assert.match(footer,/print-config-print/);assert.match(footer,/print-config-close/);
    const css=await readFile(new URL('../public/css/print-config.css',import.meta.url),'utf8');
    assert.match(css,/\.print-config-settings[^}]*overflow-y: auto/s);assert.match(css,/\.print-config-footer[^}]*flex: none/s);
});
test('Imprimir conserva parâmetros originais e ignora a página navegada na prévia',async()=>{
    const ui=Object.create(JanelaConfiguracaoImpressao.prototype);let received;
    ui.busy=false;ui.model={state:{windowsPrint:true,session:'session',revision:2,driverScale:100}};
    ui.fonte=()=>({pages:3});ui.documentPreview={source:{page:2,pages:3}};
    ui.controls=()=>{};ui.selectedScale=()=>87.5;ui.get=()=>({textContent:''});
    ui.imprimir=async(state,scale)=>{received={state,scale};return {submitted:true,printer:'Virtual',jobId:1,pages:3};};
    await ui.print();assert.deepEqual(received,{state:ui.model.state,scale:87.5});assert.equal(ui.busy,false);
});
test('Electron permite somente os dois novos módulos de apresentação previstos',()=>{
    for(const name of ['desktop-print-interface','desktop-print-preview'])assert.ok(assetPath(process.cwd(),'app://4printshops/js/app/'+name+'.js'));
    assert.equal(assetPath(process.cwd(),'app://4printshops/js/app/arbitrary.js'),null);
    assert.equal(assetPath(process.cwd(),'app://4printshops/js/app/desktop-print-preview.js.exe'),null);
});
