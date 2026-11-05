// Artefatos locais sintéticos, nunca documentos de clientes. Execute na raiz do checkout.
import {mkdir,writeFile} from 'node:fs/promises';
import path from 'node:path';
import {criarOrigemLivreto} from '../tests/helpers/pdf-booklet-fixtures.mjs';
import {gerarMontagem} from '../tests/helpers/pdf-imposition-modules.mjs';
const folder=path.resolve('.desktop-test/booklet24/examples');await mkdir(folder,{recursive:true});
for(const count of [4,8,12,16,5,7,10]){
    const source=await criarOrigemLivreto(count),r=await gerarMontagem(source,{modo:'livreto'});
    await writeFile(path.join(folder,`Origem ${count} paginas A5.pdf`),source);
    await writeFile(path.join(folder,`Livreto ${count} paginas A4 paisagem.pdf`),r.bytes);
    await writeFile(path.join(folder,`Conferencia leitura ${count} paginas.pdf`),r.bytesLeitura);
    await writeFile(path.join(folder,`Paginacao ${count}.json`),JSON.stringify(r.plan.paginacao,null,2));
}
console.log('Exemplos sintéticos em '+folder);
