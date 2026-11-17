import test from 'node:test';
import assert from 'node:assert/strict';
import {orientarFolha} from '../public/js/tools/pdf-sheet-orientation.js';

test('A: personalizada alterna repetidamente sem inverter os dados originais',()=>{
    const original=Object.freeze({id:'custom-f4',nome:'F4',largura:328,altura:478});
    let atual=original;
    for(let i=0;i<10;i++){
        atual=orientarFolha(atual.largura,atual.altura,'paisagem');assert.deepEqual(atual,{largura:478,altura:328});
        atual=orientarFolha(atual.largura,atual.altura,'retrato');assert.deepEqual(atual,{largura:328,altura:478});
    }
    assert.deepEqual(original,{id:'custom-f4',nome:'F4',largura:328,altura:478});
});
test('A: 297×420, quadrada, vírgula e medidas inválidas',()=>{
    assert.deepEqual(orientarFolha(297,420,'paisagem'),{largura:420,altura:297});
    assert.deepEqual(orientarFolha(420,297,'retrato'),{largura:297,altura:420});
    assert.deepEqual(orientarFolha(318,318,'paisagem'),{largura:318,altura:318});
    assert.deepEqual(orientarFolha('123,45','234,56','paisagem'),{largura:234.56,altura:123.45});
    for(const invalido of ['',0,-1,'inválido',Infinity])assert.equal(orientarFolha(invalido,297,'paisagem'),null);
});
