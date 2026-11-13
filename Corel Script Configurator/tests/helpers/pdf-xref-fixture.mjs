// PDF sintético literal com revisões incrementais e duas gerações do catálogo.
// Não usa save() para criar a entrada que exercita a limitação da pdf-lib.
export function criarPDFGeracoes({ concorrentes = true } = {}) {
    let texto = '%PDF-1.3\n', offsets = [0];
    const objetos = [
        '<< /Producer (Fixture de XRef) >>', '<< >>', '<< /Type /Catalog /Pages 4 0 R >>',
        '<< /Type /Pages /Kids [5 0 R 7 0 R] /Count 2 >>',
        '<< /Type /Page /Parent 4 0 R /MediaBox [0 0 200 100] /Resources << >> /Contents 6 0 R >>',
        '<< /Length 29 >>\nstream\n0 0 1 rg 10 10 20 20 re f\nendstream',
        '<< /Type /Page /Parent 4 0 R /MediaBox [0 0 100 200] /Resources << >> /Contents 8 0 R >>',
        '<< /Length 4 >>\nstream\nq Q\nendstream'
    ];
    // Comprimento do stream deve descrever exatamente seus bytes.
    const conteudo = '0 0 1 rg 10 10 20 20 re f\n';
    objetos[5] = `<< /Length ${conteudo.length} >>\nstream\n${conteudo}endstream`;
    for (let i = 1; i <= objetos.length; i++) { offsets[i] = texto.length; texto += `${i} ${i === 3 && !concorrentes ? 1 : 0} obj\n${objetos[i - 1]}\nendobj\n`; }
    const primeira = texto.length;
    texto += 'xref\n0 9\n0000000000 65535 f \n';
    for (let i = 1; i <= objetos.length; i++) texto += `${String(offsets[i]).padStart(10, '0')} ${i === 3 && !concorrentes ? '00001' : '00000'} n \n`;
    texto += `trailer\n<< /Size 9 /Root 3 ${concorrentes ? 0 : 1} R /Info 1 0 R >>\nstartxref\n${primeira}\n%%EOF\n`;
    if (concorrentes) {
        const atual = texto.length; texto += '3 1 obj\n<< /Type /Catalog /Pages 4 0 R >>\nendobj\n';
        const segunda = texto.length;
        texto += `xref\n3 1\n${String(atual).padStart(10, '0')} 00001 n \ntrailer\n<< /Size 9 /Root 3 1 R /Prev ${primeira} >>\nstartxref\n${segunda}\n%%EOF\n`;
    }
    return new TextEncoder().encode(texto);
}
