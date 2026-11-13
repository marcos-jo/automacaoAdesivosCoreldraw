import assert from 'node:assert/strict';
import {PDFDocument,PDFName,PDFHexString} from '../../public/vendor/pdf-lib/pdf-lib.js';
import {criarPDFAssinado} from './pdf-signature-fixture.mjs';

// PDF-008: substitui expectativas antigas de reescrita/remoção de assinaturas.
// Somente evidência estrutural sintética; não valida certificados ou CMS.
export async function testarAssinaturasBrowser({avaliar,pronto,esperarTela,imagem}) {
 const importar=async(bytes,name)=>{
  await avaliar(`(()=>{const dt=new DataTransfer();dt.items.add(new File([Uint8Array.from(atob(${JSON.stringify(Buffer.from(bytes).toString('base64'))}),c=>c.charCodeAt(0))],${JSON.stringify(name)},{type:'application/pdf'}));const i=document.querySelector('#pdf-arquivo');i.files=dt.files;i.dispatchEvent(new Event('change',{bubbles:true}));})()`);
  await esperarTela(`!document.querySelector('#pdf-arquivo').disabled || document.querySelector('#pdf-confirmar-descarte').open`,'Importação');
  if(await avaliar(`document.querySelector('#pdf-confirmar-descarte').open`))await avaliar(`document.querySelector('#pdf-descartar-sair').click()`);
  await pronto();
 };
 const click=s=>avaliar(`document.querySelector(${JSON.stringify(s)}).click()`);
 const status=()=>avaliar(`document.querySelector('#pdf-status').textContent`);
 await avaliar(`(()=>{window.pdf008Downloads=0;window.pdf008Anchor=HTMLAnchorElement.prototype.click;HTMLAnchorElement.prototype.click=function(){if(this.download)pdf008Downloads++;else return pdf008Anchor.call(this);};})()`);
 const perms=await PDFDocument.create();perms.addPage([200,300]);perms.catalog.set(PDFName.of('Perms'),perms.context.register(perms.context.obj({DocMDP:perms.context.register(perms.context.obj({Type:'Sig',ByteRange:[0,10,20,10],Contents:PDFHexString.of('01020304')}))})));
 try {
  for(const [name,bytes] of [['aplicada',await criarPDFAssinado()],['multipla',await criarPDFAssinado({paginas:[1,3]})],['perms',await perms.save()]]){
   await importar(bytes,name+'.pdf');assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-alerta').hidden`),false);
   for(const selector of ['#pdf-girar-direita','#pdf-recortar','#pdf-salvar-copia','#pdf-assinaturas-resolver']){
    await click(selector);if(selector==='#pdf-assinaturas-resolver'){assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-expandido').hidden`),false);assert.equal(await avaliar(`document.querySelector('#pdf-impressao-dialogo').open`),false);}else{await esperarTela(`/assinatura digital/.test(document.querySelector('#pdf-status').textContent) && !document.querySelector('#pdf-arquivo').disabled`,'Bloqueio '+selector);assert.match(await status(),/bloqueado/);}
   }
   assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),true);assert.equal(await avaliar(`!!document.querySelector('#app-exportacao-dialogo')?.open`),false);
   await imagem('pdf008-'+name);
  }
  // Worker real da conferência: origem comum mantida, assinada removível.
  await click('#pdf-uniao-abrir');
  const signed=Buffer.from(await criarPDFAssinado()).toString('base64');
  await avaliar(`(async()=>{const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');const p=await PDFDocument.create();p.addPage([200,300]);const dt=new DataTransfer();dt.items.add(new File([await p.save()],'comum.pdf',{type:'application/pdf'}));dt.items.add(new File([Uint8Array.from(atob(${JSON.stringify(signed)}),c=>c.charCodeAt(0))],'assinada.pdf',{type:'application/pdf'}));const i=document.querySelector('#pdf-uniao-arquivos');i.files=dt.files;i.dispatchEvent(new Event('change',{bubbles:true}));})()`);
  await esperarTela(`!!document.querySelector('#pdf-uniao-preparacao li[data-falha="true"] button:not([hidden])')`,'Assinada recusada no worker');
  assert.match(await avaliar(`document.querySelector('#pdf-uniao-preparacao li[data-falha="true"]').textContent`),/assinatura digital/);assert.equal(await avaliar(`document.querySelector('#pdf-uniao-gerar').disabled`),true);
  await imagem('pdf008-uniao-recusada');await click('#pdf-uniao-preparacao li[data-falha="true"] button[aria-label^="Remover"]');await esperarTela(`!document.querySelector('#pdf-uniao-gerar').disabled`,'Comum mantido');await click('#pdf-uniao-fechar');
  // Entradas de janelas reais, sem alterar seus controles ou gerar montagem.
  const montage=await avaliar(`(async()=>{const {PDFDocument,PDFName,PDFHexString}=await import('/vendor/pdf-lib/pdf-lib.js');const {PDFWorkspace}=await import('/js/tools/pdf-workspace.js');const p=await PDFDocument.create();p.addPage([200,300]);p.catalog.set(PDFName.of('Perms'),p.context.obj({DocMDP:p.context.obj({Type:'Sig',ByteRange:[0,10,20,10],Contents:PDFHexString.of('010203')})}));const b=await p.save(),ws=new PDFWorkspace();await ws.abrir({name:'perms.pdf',arrayBuffer:async()=>b.slice().buffer});const {JanelaMontagem}=await import('/js/tools/pdf-imposition-ui.js'),{JanelaLivreto}=await import('/js/tools/pdf-booklet-ui.js');const results=[];for(const Class of [JanelaMontagem,JanelaLivreto]){const root=document.createElement('div');document.body.append(root);const ui=new Class(root,{snapshot:()=>ws.obterBytes()});try{await ui.open();results.push({message:ui.get('status').textContent,hasResult:!!ui.result});}finally{await ui.destroy();root.remove();}}return results;})()`);
  assert.equal(montage.length,2);for(const r of montage){assert.match(r.message,/assinatura digital/);assert.equal(r.hasResult,false);}
  await importar(await criarPDFAssinado({vazia:true}),'campo-vazio.pdf');assert.match(await avaliar(`document.querySelector('#pdf-assinaturas-titulo').textContent`),/Campo|campo/);await click('#pdf-girar-direita');await pronto();assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),false);
  assert.equal(await avaliar('pdf008Downloads'),0);
  await importar(await criarPDFAssinado({paginas:[]}),'comum-final.pdf');await click('#pdf-girar-direita');await pronto();assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-alerta').hidden`),true);
  console.log('PASSOU PDF-008 Chrome: visualização, rotação/recorte/exportação bloqueados e Ver restrições sem preparo, worker de união, montagem/livreto sem resultado, campo vazio e comum editáveis. Sem downloads assinados.');
 } finally {await avaliar('HTMLAnchorElement.prototype.click=pdf008Anchor');}
}
