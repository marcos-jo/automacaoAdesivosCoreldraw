import assert from 'node:assert/strict';
import {writeFile} from 'node:fs/promises';import path from 'node:path';

// Opera os controles reais: recolher não pode mudar resultado/consentimento.
export async function testarInterfaceBasicaMontagem({ev,cdp,wait,screenshot,folder}){
    const ready=()=>wait(()=>ev('!document.querySelector(".pdf-montagem [data-id=salvar]").disabled'),'montagem básica pronta');
    const state=()=>ev(`(()=>{const d=document.querySelector('.pdf-montagem'),f=id=>d.querySelector('[data-id="'+id+'"]'),visible=id=>f(id).checkVisibility();return {advanced:f('avancadas').open,basic:[...d.querySelectorAll('form input,form select')].filter(e=>e.checkVisibility()).map(e=>e.dataset.id),quantity:visible('quantidade'),custom:visible('largura')&&visible('altura'),mode:f('modo').value,summary:f('distribuicao').textContent,status:f('status').textContent,canvas:d.querySelector('canvas').width,save:f('salvar').disabled,print:f('imprimir').disabled,grade:f('grade').value,quantityValue:f('quantidade').value,pages:f('paginas').value,dimensions:[f('largura').value,f('altura').value]};})()`);
    const edit=async(id,value)=>{await ev(`(()=>{const f=document.querySelector('.pdf-montagem [data-id=${id}]');f.value=${JSON.stringify(value)};f.dispatchEvent(new Event('input',{bubbles:true}));})()`);await ready();};
    const toggle=()=>ev('document.querySelector(".pdf-montagem [data-id=avancadas] > summary").click()');
    await ready();const before=await state();assert.equal(before.advanced,false);assert.deepEqual(before.basic,['modo','paginas','formato','orientacao','quantidade']);assert.match(before.summary,/Distribuição automática/);assert.ok(before.canvas>0);await screenshot('interface-basica-repeat');
    assert.equal(await ev("document.querySelector('.pdf-montagem [data-id=medidas-resumo]').checkVisibility()"),true);assert.match(await ev("document.querySelector('.pdf-montagem [data-id=medidas-resumo]').textContent"),/escala efetiva/);
    await toggle();const expanded=await state();assert.equal(expanded.advanced,true);for(const id of ['grade','rotacao','escalaModo','pecaLargura','esquerda','espacoX','sangria','modoMarcas','guias'])assert.ok(expanded.basic.includes(id),id+' deve continuar acessível');assert.equal(expanded.status,before.status);assert.equal(expanded.canvas,before.canvas);
    // O disclosure nativo também responde pelo teclado.
    await ev('document.querySelector(".pdf-montagem [data-id=avancadas] > summary").focus()');await cdp('Input.dispatchKeyEvent',{type:'rawKeyDown',key:' ',code:'Space',windowsVirtualKeyCode:32});await cdp('Input.dispatchKeyEvent',{type:'keyUp',key:' ',code:'Space',windowsVirtualKeyCode:32});await wait(()=>ev('!document.querySelector(".pdf-montagem [data-id=avancadas]").open'),'recolher pelo teclado');assert.equal((await state()).advanced,false);
    await edit('modo','nup');const nup=await state();assert.deepEqual(nup.basic,['modo','paginas','formato','orientacao','nupDistribuicao']);assert.equal(nup.quantity,false);assert.equal(nup.grade,'4');assert.match(nup.summary,/página\(s\) por folha/);await screenshot('interface-basica-nup');
    await toggle();const nupAdvanced=await state();assert.equal(nupAdvanced.basic.includes('pecaLargura'),false);assert.equal(nupAdvanced.quantity,false);await toggle();
    await edit('modo','repeat');assert.equal((await state()).quantityValue,before.quantityValue);await edit('grade','auto');
    await edit('formato','custom');const custom=await state();assert.equal(custom.custom,true);assert.equal(custom.advanced,false);await screenshot('interface-basica-folha-personalizada');
    await edit('largura','320');await edit('altura','450');assert.deepEqual((await state()).dimensions,['320','450']);await edit('formato','A3');assert.equal((await state()).custom,false);assert.deepEqual((await state()).dimensions,['297','420']);
    await toggle();await edit('modoMarcas','grupo');const changed=await state();await screenshot('interface-avancada');await toggle();assert.equal((await state()).status,changed.status);assert.equal((await state()).save,false);await toggle();await edit('modoMarcas','individuais');await toggle();
    assert.equal(await ev('document.querySelector("#pdf-alteracoes-pendentes").hidden'),false);
    await writeFile(path.join(folder,'interface-basica.json'),JSON.stringify({repeat:before,nup,custom:custom.basic,keyboard:true,collapseDoesNotChangeResult:true,advancedValuesPreserved:true,originalPending:true},null,2));
}
