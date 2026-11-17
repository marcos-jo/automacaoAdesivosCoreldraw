import { PDFFormatos } from './pdf-formats.js';

// Uma carga e uma fila de gravação para todos os consumidores do renderer.
let compartilhado;
export function obterCatalogoFormatos() {
    return compartilhado ??= new PDFFormatos();
}
