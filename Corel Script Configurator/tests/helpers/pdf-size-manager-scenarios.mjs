import assert from 'node:assert/strict';

const key='4printshops.pdf.formatos.v1';
export const seed={versao:1,extensao:{preservar:true},formatos:[
    {id:'custom-sra3',nome:'SRA3',largura:320,altura:450,origem:'sintetico'},
    ...[['F4',328,478],['F6',318,318],['F8',238,328]].map(([nome,largura,altura])=>({id:'custom-'+nome,nome,largura,altura,origem:'sintetico'})),{id:'A3',nome:'A3',largura:297,altura:420},{id:'A5',nome:'A5',largura:148,altura:210}]};

// Exercita eventos reais das UIs. Somente catálogos e PDFs sintéticos no perfil isolado.
export async function preparar(ev,wait) {
    await ev(`document.querySelector('[data-tool=pdfs]').click()`);
    await wait(()=>ev(`!!document.querySelector('#pdf-redim-formato')`),'Redimensionar');
    await ev(`(async()=>{
        const{obterCatalogoFormatos}=await import('/js/tools/pdf-format-catalog.js');window.catalog=obterCatalogoFormatos();await catalog.pronta;
        const{PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');const p=await PDFDocument.create();p.addPage([255.118,141.732]).drawText('SINTETICO',{x:10,y:50,size:10});window.sourceBytes=await p.save();
        const{JanelaMontagem}=await import('/js/tools/pdf-imposition-ui.js');
        window.repeat=new JanelaMontagem(document.querySelector('#app'),{snapshot:async()=>sourceBytes.slice(),imprimir:()=>{throw Error('Impressão proibida no teste');}});
        window.nup=new JanelaMontagem(document.querySelector('#app'),{snapshot:async()=>sourceBytes.slice(),imprimir:()=>{throw Error('Impressão proibida no teste');}});
        nup.get('modo').value='nup';nup.get('modo').dispatchEvent(new Event('input',{bubbles:true}));
        window.choose=(el,value,type='change')=>{el.value=String(value);el.dispatchEvent(new Event(type,{bubbles:true}));};
        window.fields=()=>Object.fromEntries(['nome','largura','altura'].map(k=>[k,document.querySelector('#pdf-tamanhos [data-campo='+k+']')]));
        window.lists=()=>[...document.querySelectorAll('#pdf-redim-formato,dialog.pdf-montagem [data-id=formato]')].map(e=>[...e.options].map(o=>o.value));
        window.openManager=()=>{const b=document.querySelector('[data-command=gerenciar-tamanhos]');b.closest('details').open=true;b.click();};
        window.confirmations=[];window.allowDelete=true;window.confirm=message=>{confirmations.push(message);return allowDelete;};
        await repeat.open();await nup.open();
    })()`);
    await wait(()=>ev('!!repeat.result&&!repeat.busy&&!repeat.dirty'),'prévia inicial');
}
export async function executar(ev,wait) {
    const checks={A:[],B:[]};
    const raw=()=>ev(`localStorage.getItem('${key}')`);
    // Selecionar nas duas montagens abertas/instanciadas e no cadastro rápido.
    await ev(`choose(repeat.get('formato'),'custom-sra3','input');choose(nup.get('formato'),'custom-sra3','input');choose(document.querySelector('#pdf-redim-formato'),'custom-sra3');`);
    await wait(()=>ev('!repeat.busy&&!repeat.dirty&&!nup.busy&&!nup.dirty'),'prévia SRA3');
    const dimensions=await ev(`[repeat,nup].map(u=>[u.get('largura').value,u.get('altura').value])`);
    const original=await raw();assert.equal(original,JSON.stringify(seed));
    await ev(`document.querySelector('#pdf-formato-excluir').click()`);
    await wait(()=>ev(`!catalog.listar().some(f=>f.nome==='SRA3')`),'excluir SRA3 no Redimensionar');
    assert.equal(await ev(`lists().every(l=>!l.includes('SRA3')&&!l.includes('custom-sra3'))`),true);
    assert.deepEqual(await ev(`[repeat,nup].map(u=>[u.get('largura').value,u.get('altura').value])`),dimensions);
    assert.deepEqual(await ev(`[repeat.get('formato').value,nup.get('formato').value,document.querySelector('#pdf-redim-formato').value]`),['custom','custom','personalizado']);
    checks.A.push('Exclusão no Redimensionar atualiza ambas as montagens e preserva geometria/PDF');
    await ev(`(async()=>{await repeat.close();openManager()})()`);
    await wait(()=>ev(`document.querySelector('#pdf-tamanhos').open`),'menu Arquivo → Gerenciar tamanhos');
    const fill=async(nome,w,h)=>ev(`(()=>{const f=fields();f.nome.value=${JSON.stringify(nome)};f.largura.value=${JSON.stringify(String(w))};f.altura.value=${JSON.stringify(String(h))};document.querySelector('#pdf-tamanhos form').requestSubmit();})()`);
    await fill('TAMANHO TESTE GERENCIADOR',140,180);
    await wait(()=>ev(`catalog.listar().some(f=>f.nome==='TAMANHO TESTE GERENCIADOR')`),'adicionar pelo gerenciador');
    const id=await ev(`catalog.listar().find(f=>f.nome==='TAMANHO TESTE GERENCIADOR').id`);
    assert.equal(await ev(`lists().every(l=>l.includes(${JSON.stringify(id)}))`),true);
    // Duplicata canônica recusada sem gravação.
    await ev(`document.querySelector('#pdf-tamanhos [data-campo=novo]').click()`);
    const beforeDuplicate=await raw();await fill(' tamanho   teste gerenciador ',99,199);
    await wait(()=>ev(`document.querySelector('#pdf-tamanhos [data-campo=status]').textContent.includes('existe')`),'nome duplicado');assert.equal(await raw(),beforeDuplicate);
    await ev(`choose(document.querySelector('#pdf-tamanhos [data-campo=lista]'),${JSON.stringify(id)})`);
    await fill('TAMANHO TESTE EDITADO',150,190);
    await wait(()=>ev(`catalog.listar().some(f=>f.id===${JSON.stringify(id)}&&f.largura===150)`),'editar sem mudar ID');
    assert.equal(await ev(`lists().every(l=>l.includes(${JSON.stringify(id)}))`),true);
    const edited=JSON.parse(await raw()).formatos.find(f=>f.id===id);assert.deepEqual([edited.nome,edited.largura,edited.altura],['TAMANHO TESTE EDITADO',150,190]);
    checks.B.push('Menu global adiciona/edita; mesmos ID e registros nas três ferramentas; nome duplicado recusado');
    // Seleção em montagem é preservada como geometria avulsa ao excluir pelo gerenciador.
    const beforeCancel=await raw();
    await ev(`choose(repeat.get('formato'),${JSON.stringify(id)},'input');choose(nup.get('formato'),${JSON.stringify(id)},'input');allowDelete=false;document.querySelector('#pdf-tamanhos [data-campo=excluir]').click()`);
    await wait(()=>ev(`!document.querySelector('#pdf-tamanhos [data-campo=excluir]').disabled`),'cancelar exclusão');assert.equal(await raw(),beforeCancel);assert.equal(await ev(`catalog.listar().some(f=>f.id===${JSON.stringify(id)})`),true);
    await ev(`allowDelete=true;document.querySelector('#pdf-tamanhos [data-campo=excluir]').click()`);
    await wait(()=>ev(`!catalog.listar().some(f=>f.id===${JSON.stringify(id)})`),'excluir pelo gerenciador');
    assert.equal(await ev(`lists().every(l=>!l.includes(${JSON.stringify(id)}))`),true);assert.deepEqual(await ev(`[repeat,nup].map(u=>[u.get('formato').value,Number(u.get('largura').value),Number(u.get('altura').value)])`),[['custom',150,190],['custom',150,190]]);
    assert.equal(await ev(`confirmations.filter(m=>m.includes('Excluir')).length`),3);
    checks.B.push('Confirmação cancelada preserva registro; confirmação aceita remove seleção sem aplicar operação');
    // Editar e excluir tamanhos iniciais é permitido e não os recria.
    await ev(`choose(document.querySelector('#pdf-tamanhos [data-campo=lista]'),'A3')`);
    await fill('A3 SINTETICO EDITADO',300,430);await wait(()=>ev(`catalog.listar().some(f=>f.id==='A3'&&f.largura===300)`),'editar inicial');
    await ev(`choose(document.querySelector('#pdf-tamanhos [data-campo=lista]'),'A5');document.querySelector('#pdf-tamanhos [data-campo=excluir]').click()`);
    await wait(()=>ev(`!catalog.listar().some(f=>f.id==='A5')`),'excluir inicial A5');
    // Formulário desatualizado não sobrescreve uma alteração recebida.
    await ev(`(async()=>{choose(document.querySelector('#pdf-tamanhos [data-campo=lista]'),'custom-F8');await catalog.salvar({nome:'F8 SINTETICO EXTERNO',largura:238,altura:328},'custom-F8');document.querySelector('#pdf-tamanhos form').requestSubmit()})()`);
    await wait(()=>ev(`document.querySelector('#pdf-tamanhos [data-campo=status]').textContent.includes('outra ferramenta')`),'edição obsoleta recusada');
    await ev(`(async()=>{await catalog.salvar({nome:'F8',largura:238,altura:328},'custom-F8');document.querySelector('#pdf-tamanhos [data-campo=novo]').click()})()`);
    await fill('PERSISTENCIA SINTETICA',123,234);await wait(()=>ev(`catalog.listar().some(f=>f.nome==='PERSISTENCIA SINTETICA')`),'formato para reabertura');
    const persistent=await ev(`catalog.listar().find(f=>f.nome==='PERSISTENCIA SINTETICA').id`);
    await ev(`(async()=>{document.querySelector('#pdf-tamanhos [data-campo=fechar]').click();await repeat.open();choose(repeat.get('formato'),'A3','input');choose(repeat.get('rotacao'),'90','input')})()`);
    for(const [o,w,h]of [['paisagem',430,300],['retrato',300,430]]){
        await ev(`choose(repeat.get('orientacao'),'${o}','input')`);await wait(()=>ev('!!repeat.result&&!repeat.busy&&!repeat.dirty'),'prévia '+o);
        assert.deepEqual(await ev(`[repeat.result.plan.folha.largura,repeat.result.plan.folha.altura,repeat.config().rotacao]`),[w,h,90]);
    }
    assert.equal(await ev('repeat.originalBytes.every((b,i)=>b===sourceBytes[i])'),true);
    await ev(`(async()=>{await repeat.close();await nup.close()})()`);
    const persisted=await raw();assert.deepEqual(JSON.parse(persisted).formatos.filter(f=>/^custom-F/.test(f.id)),seed.formatos.filter(f=>/^custom-F/.test(f.id)));
    checks.A.push('A3 editável e A5 excluível; orientação e arte independentes; outros registros/IDs/metadados preservados');
    checks.B.push('Edição obsoleta recusada; cadastro persistente preparado para reinício');
    return {checks,persisted,persistent,deleted:id};
}
export async function conferirReabertura(ev,wait,result) {
    assert.equal(await ev(`localStorage.getItem('${key}')`),result.persisted);
    await preparar(ev,wait);
    assert.equal(await ev(`lists().every(l=>l.includes(${JSON.stringify(result.persistent)})&&!l.includes('custom-sra3')&&!l.includes('SRA3')&&!l.includes('A5')&&!l.includes(${JSON.stringify(result.deleted)}))`),true);
    await ev(`(async()=>{await repeat.close();await nup.close();openManager()})()`);await wait(()=>ev(`document.querySelector('#pdf-tamanhos').open`),'gerenciador reaberto');
    assert.equal(await ev(`document.querySelector('#pdf-tamanhos [data-campo=lista] option[value="${result.persistent}"]')?.textContent.includes('PERSISTENCIA SINTETICA')`),true);
    assert.equal(await ev(`new Set(catalog.listar().map(f=>f.id)).size===catalog.listar().length`),true);
    await ev(`document.querySelector('#pdf-tamanhos [data-campo=fechar]').click()`);
}
