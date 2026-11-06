import test from "node:test";
import assert from "node:assert/strict";
import { setTimeout as aguardar } from "node:timers/promises";
import { Elemento } from "./helpers/pdf-dom.mjs";
import { PDFControlesVisualizacao } from "../public/js/tools/pdf-view-controls.js";

function iniciar() {
    const elementos = new Map(["visualizacao", "zoom-valor", "zoom-pagina", "zoom-largura", "zoom-real", "zoom-mais", "zoom-menos"].map(id => ["#pdf-" + id, new Elemento("div")]));
    const container = elementos.get("#pdf-visualizacao");
    container.clientWidth = 632; container.clientHeight = 432;
    const previa = {}, estado = { chamadas: 0, permitido: true };
    const controles = new PDFControlesVisualizacao(id => elementos.get(id), previa, async () => { estado.chamadas++; }, () => estado.permitido);
    return { elementos, container, previa, estado, controles };
}
test("ajuste usa largura e altura disponíveis; 100% e zoom manual não alteram o PDF", () => {
    const { previa, controles } = iniciar();
    try {
        assert.equal(previa.calcularEscala({ width: 300, height: 800 }), .5);
        assert.equal(previa.calcularEscala({ width: 1200, height: 300 }), .5);
        controles.modo = "largura"; assert.equal(previa.calcularEscala({ width: 300, height: 800 }), 2);
        controles.modo = "real"; controles.zoom = 1; assert.equal(previa.calcularEscala({ width: 300, height: 800 }), 1);
        controles.modo = "manual"; controles.zoom = 6;
        assert.ok(previa.calcularEscala({ width: 5000, height: 5000 }) ** 2 * 25000000 <= 16000000 + 1e-6);
    } finally { controles.destruir(); }
});
test("roda somente no centro, agrupamento de eventos e liberação de listeners/timer", async () => {
    const { container, previa, controles, estado } = iniciar();
    for (let i = 0; i < 10; i++) assert.equal(container.emitir("wheel", container, { deltaY: -30, deltaMode: 0, clientX: 20, clientY: 20 }).defaultPrevented, true);
    await aguardar(100); assert.equal(estado.chamadas, 1);
    assert.equal(controles.modo, "manual");
    estado.permitido = false;
    assert.equal(container.emitir("wheel", container, { deltaY: 50 }).defaultPrevented, false);
    estado.permitido = true; controles.ampliar(2); controles.destruir(); await aguardar(100);
    assert.equal(estado.chamadas, 1); assert.equal(previa.calcularEscala, undefined);
    assert.ok([...container.listeners.values()].every(lista => lista.length === 0));
});

// Modelo de layout do navegador: padding16, canvas centralizado horizontalmente
// quando cabe, e scroll limitado à extensão do conteúdo. Não há offset de toolbar.
function comLayout(clientWidth,clientHeight){
    const {container,previa,controles,estado}=iniciar();container.clientWidth=clientWidth;container.clientHeight=clientHeight;
    let width=900,height=600,left=0,top=0;
    Object.defineProperties(container,{scrollWidth:{get:()=>Math.max(clientWidth,width+32)},scrollHeight:{get:()=>Math.max(clientHeight,height+32)},scrollLeft:{get:()=>left,set:v=>{left=Math.max(0,Math.min(container.scrollWidth-clientWidth,v));}},scrollTop:{get:()=>top,set:v=>{top=Math.max(0,Math.min(container.scrollHeight-clientHeight,v));}}});
    const canvas={getBoundingClientRect:()=>({left:174+16+Math.max(0,(clientWidth-32-width)/2)-left,top:195+16-top,width,height})};
    container.getBoundingClientRect=()=>({left:174,top:195,width:clientWidth+10,height:clientHeight+10});container.querySelector=()=>canvas;
    controles.zoom=1;controles.atualizar=async()=>{estado.chamadas++;const scale=previa.calcularEscala({width:900,height:600});width=Math.ceil(900*scale);height=Math.ceil(600*scale);container.scrollLeft=left;container.scrollTop=top;};
    return {container,controles,canvas,estado};
}
test('zoom preserva ponto quando há rolagem suficiente',async()=>{
    const {container,controles,canvas}=comLayout(928,549),point={x:643,y:474.5},before=canvas.getBoundingClientRect(),fx=(point.x-before.left)/before.width,fy=(point.y-before.top)/before.height;
    try{controles.ampliar(Math.exp(.36),point);await aguardar(100);const r=canvas.getBoundingClientRect();assert.ok(Math.abs(r.left+fx*r.width-point.x)<1);assert.ok(Math.abs(r.top+fy*r.height-point.y)<1);assert.ok(container.scrollTop>0);}finally{controles.destruir();}
});
test('53,45 px são o residual do limite: necessidade164,45 contra máximo111',async()=>{
    const {container,controles,canvas}=comLayout(1002,781),point={x:680,y:590.5},before=canvas.getBoundingClientRect(),fy=(point.y-before.top)/before.height;
    try{controles.ampliar(Math.exp(.36),point);await aguardar(100);const r=canvas.getBoundingClientRect(),required=container.scrollTop+r.top+fy*r.height-point.y;assert.ok(Math.abs(required-164.45)<.001);assert.equal(container.scrollHeight-container.clientHeight,111);assert.equal(container.scrollTop,111);assert.ok(Math.abs(r.top+fy*r.height-point.y-(required-111))<.001);}finally{controles.destruir();}
});
test('zoom pelo controle de interface conserva o centro, inclusive com scroll prévio',async()=>{
    const {container,controles,canvas}=comLayout(632,432);container.scrollLeft=120;container.scrollTop=80;
    const point={x:174+316,y:195+216},before=canvas.getBoundingClientRect(),fx=(point.x-before.left)/before.width,fy=(point.y-before.top)/before.height;
    try{controles.ampliar(1.2);await aguardar(100);const r=canvas.getBoundingClientRect();assert.ok(Math.abs(r.left+fx*r.width-point.x)<1);assert.ok(Math.abs(r.top+fy*r.height-point.y)<1);}finally{controles.destruir();}
});
