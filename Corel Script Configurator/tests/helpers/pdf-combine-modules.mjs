import { readFile } from 'node:fs/promises';
const library = new URL('../../public/vendor/pdf-lib/pdf-lib.js', import.meta.url).href;
async function load(name, replacements = {}) {
    let code = await readFile(new URL('../../public/js/tools/' + name, import.meta.url), 'utf8');
    for (const [from, to] of Object.entries({ '/vendor/pdf-lib/pdf-lib.js': library, ...replacements })) code = code.replaceAll(JSON.stringify(from), JSON.stringify(to));
    return 'data:text/javascript;base64,' + Buffer.from(code).toString('base64');
}
export const imageURL = await load('pdf-image-import.js');
export const combineURL = await load('pdf-combine-session.js', { './pdf-image-import.js': imageURL });
export const { PDFCombinacao } = await import(combineURL);
export const { analisarImagem, importarImagem, matrizImagem } = await import(imageURL);
export const mergeModule = previewURL => load('pdf-merge.js', { './pdf-combine-session.js': combineURL, './pdf-preview.js': previewURL, './pdf-import-review.js':new URL('../../public/js/tools/pdf-import-review.js',import.meta.url).href });
