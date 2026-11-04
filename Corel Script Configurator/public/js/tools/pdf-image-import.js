import { PDFDocument, pushGraphicsState, popGraphicsState, concatTransformationMatrix, drawObject } from "/vendor/pdf-lib/pdf-lib.js";

export const LIMITE_COMBINACAO = 512 * 1024 * 1024;
const texto = bytes => String.fromCharCode(...bytes);

function orientacaoExif(bytes) {
    if (texto(bytes.subarray(0, 6)) !== "Exif\0\0") return 1;
    const view = new DataView(bytes.buffer, bytes.byteOffset + 6, bytes.length - 6);
    const little = view.getUint16(0) === 0x4949;
    if (view.getUint16(2, little) !== 42) throw new Error("Metadados EXIF inválidos.");
    const offset = view.getUint32(4, little);
    const count = view.getUint16(offset, little);
    for (let i = 0; i < count; i++) {
        const entry = offset + 2 + i * 12;
        if (view.getUint16(entry, little) !== 0x112) continue;
        if (view.getUint16(entry + 2, little) !== 3 || view.getUint32(entry + 4, little) !== 1) throw new Error("Orientação EXIF inválida.");
        const value = view.getUint16(entry + 8, little);
        if (value < 1 || value > 8) throw new Error("Orientação EXIF não suportada.");
        return value;
    }
    return 1;
}

export function analisarImagem(bytes, extensao) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let width, height, orientation = 1;
    if (extensao === "jpg" || extensao === "jpeg") {
        if (view.getUint16(0) !== 0xffd8) throw new Error("Conteúdo JPEG inválido.");
        let pos = 2;
        while (pos + 4 <= bytes.length) {
            if (bytes[pos++] !== 255) throw new Error("Marcador JPEG inválido.");
            while (bytes[pos] === 255) pos++;
            const marker = bytes[pos++];
            if (marker === 0xda || marker === 0xd9) break;
            const length = view.getUint16(pos);
            if (length < 2 || pos + length > bytes.length) throw new Error("Segmento JPEG truncado.");
            const data = bytes.subarray(pos + 2, pos + length);
            if (marker === 0xe2 && texto(data.subarray(0, 12)) === "ICC_PROFILE\0") throw new Error("JPEG com perfil ICC. Converta em PDF com gerenciamento de cor no aplicativo de origem antes de combinar.");
            if (marker === 0xe1 && texto(data.subarray(0, 6)) === "Exif\0\0") orientation = orientacaoExif(data);
            if ([0xc0, 0xc1, 0xc2].includes(marker)) {
                if (data[0] !== 8 || ![1, 3].includes(data[5])) throw new Error("JPEG CMYK ou codificação não suportada. Prepare um PDF no aplicativo de origem para preservar as cores.");
                height = view.getUint16(pos + 3);
                width = view.getUint16(pos + 5);
            }
            pos += length;
        }
    } else if (extensao === "png") {
        if (texto(bytes.subarray(0, 8)) !== "\x89PNG\r\n\x1a\n") throw new Error("Conteúdo PNG inválido.");
        let pos = 8, srgb = false, chroma = false;
        while (pos + 12 <= bytes.length) {
            const length = view.getUint32(pos), type = texto(bytes.subarray(pos + 4, pos + 8));
            if (length > bytes.length - pos - 12) throw new Error("PNG truncado.");
            if (type === "IHDR") {
                width = view.getUint32(pos + 8); height = view.getUint32(pos + 12);
                if (bytes[pos + 16] === 16) throw new Error("PNG de 16 bits. Prepare um PDF para preservar a profundidade de cor.");
            }
            if (type === "iCCP") throw new Error("PNG com perfil ICC. Prepare um PDF com gerenciamento de cor antes de combinar.");
            if (type === "sRGB") srgb = true;
            if (type === "cHRM") chroma = true;
            if (type === "gAMA" && (length !== 4 || view.getUint32(pos + 8) !== 45455)) throw new Error("PNG com gama não padrão. Prepare um PDF no aplicativo de origem.");
            if (type === "acTL") throw new Error("PNG animado não suportado.");
            if (type === "eXIf") {
                const exif = new Uint8Array(length + 6);
                exif.set([69, 120, 105, 102, 0, 0]); exif.set(bytes.subarray(pos + 8, pos + 8 + length), 6);
                orientation = orientacaoExif(exif);
            }
            pos += length + 12;
            if (type === "IEND") break;
        }
        if (chroma && !srgb) throw new Error("PNG com cromaticidade personalizada. Prepare um PDF para preservar as cores.");
    } else throw new Error("Selecione PDF, JPG/JPEG ou PNG.");
    if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1 || width > 20000 || height > 20000 || width * height > 40000000) throw new Error("Imagem inválida ou acima de 40 megapixels / 20.000 pixels por lado.");
    return { width, height, orientation };
}

export function matrizImagem(width, height, orientation) {
    return [null, [width, 0, 0, height, 0, 0], [-width, 0, 0, height, width, 0],
        [-width, 0, 0, -height, width, height], [width, 0, 0, -height, 0, height],
        [0, -width, -height, 0, height, width], [0, -width, height, 0, 0, width],
        [0, width, height, 0, 0, 0], [0, width, -height, 0, height, 0]][orientation];
}

export async function importarImagem(bytes, extensao) {
    const info = analisarImagem(bytes, extensao);
    const pdf = await PDFDocument.create();
    const data = new Uint8Array(bytes);
    const image = extensao === "png" ? await pdf.embedPng(data) : await pdf.embedJpg(data);
    // Dimensão inicial explícita a 150 ppi. DPI declarados pelo arquivo não definem escala.
    const width = info.width * 72 / 150, height = info.height * 72 / 150;
    const swapped = info.orientation >= 5;
    const page = pdf.addPage(swapped ? [height, width] : [width, height]);
    const name = page.node.newXObject("Image", image.ref);
    page.pushOperators(pushGraphicsState(), concatTransformationMatrix(...matrizImagem(width, height, info.orientation)), drawObject(name), popGraphicsState());
    return { pdf, info };
}
