import { createHash } from 'node:crypto';

const padding = Buffer.from('28bf4e5e4e758a4164004e56fffa01082e2e00b6d0683e802f0ca9fe6453697a', 'hex');
const digest = bytes => createHash('md5').update(bytes).digest();
const pad = password => Buffer.concat([Buffer.from(password, 'ascii'), padding]).subarray(0, 32);

function rc4(key, data) {
    const state = Array.from({ length: 256 }, (_, i) => i);
    let i = 0, j = 0;
    const swap = () => { [state[i], state[j]] = [state[j], state[i]]; };
    for (; i < 256; i++) { j = (j + state[i] + key[i % key.length]) % 256; swap(); }
    i = j = 0;
    return Buffer.from(data.map(byte => {
        i = (i + 1) % 256; j = (j + state[i]) % 256; swap();
        return byte ^ state[(state[i] + state[j]) % 256];
    }));
}

// PDF sintético sem strings/streams de conteúdo, com proteção Standard R2.
export function pdfProtegido(password = '') {
    const id = Buffer.from('0123456789abcdef0123456789abcdef', 'hex');
    const permissions = Buffer.alloc(4);
    permissions.writeInt32LE(-60); // Permite imprimir, restringe modificar/copiar.
    const owner = rc4(digest(pad('proprietario-sintetico')).subarray(0, 5), pad(password));
    const key = digest(Buffer.concat([pad(password), owner, permissions, id])).subarray(0, 5);
    const user = rc4(key, padding);
    const objects = [
        '<< /Type /Catalog /Pages 2 0 R >>',
        '<< /Type /Pages /Count 1 /Kids [3 0 R] >>',
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 400] /Resources << >> >>',
        `<< /Filter /Standard /V 1 /R 2 /O <${owner.toString('hex')}> /U <${user.toString('hex')}> /P -60 >>`
    ];
    let text = '%PDF-1.4\n';
    const offsets = objects.map((object, index) => {
        const offset = text.length;
        text += `${index + 1} 0 obj\n${object}\nendobj\n`;
        return offset;
    });
    const xref = text.length, hex = id.toString('hex');
    text += 'xref\n0 5\n0000000000 65535 f \n' + offsets.map(offset => String(offset).padStart(10, '0') + ' 00000 n \n').join('');
    return Buffer.from(text + `trailer\n<< /Size 5 /Root 1 0 R /Encrypt 4 0 R /ID [<${hex}> <${hex}>] >>\nstartxref\n${xref}\n%%EOF\n`, 'ascii');
}
