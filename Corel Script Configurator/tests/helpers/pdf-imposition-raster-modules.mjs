import {readFile} from 'node:fs/promises';
import {copiaImpressaoUrl} from './pdf-resize-modules.mjs';
import {compatibilityUrl} from './pdf-imposition-modules.mjs';
const url=file=>new URL(file,import.meta.url).href;
let code=await readFile(new URL('../../public/js/tools/pdf-imposition-raster.js',import.meta.url),'utf8');
for(const [from,to]of Object.entries({'/vendor/pdf-lib/pdf-lib.js':url('../../public/vendor/pdf-lib/pdf-lib.js'),'./pdf-print-copy.js':copiaImpressaoUrl,'./pdf-imposition-compatibility.js':compatibilityUrl,'./pdf-print-raster.js':url('../../public/js/tools/pdf-print-raster.js'),'./pdf-imposition-preview.js':url('../../public/js/tools/pdf-imposition-preview.js')}))code=code.replaceAll("'"+from+"'",JSON.stringify(to));
export const {planejarImagemMontagem,avaliarImagemMontagem,prepararImagemMontagem}=await import('data:text/javascript;base64,'+Buffer.from(code).toString('base64'));
