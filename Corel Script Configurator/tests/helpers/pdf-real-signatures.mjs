// Somente fixtures sintéticas. Não assina arquivos de clientes e não instala
// certificados/chaves. OpenSSL já instalado gera CMS detached SHA-256/RSA.
import { mkdtemp, readFile, writeFile, unlink } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { createHash } from "node:crypto";
import { PDFDocument, PDFName, PDFRef, PDFString, PDFHexString } from "../../public/vendor/pdf-lib/pdf-lib.js";
import { criarPDFAssinado } from "./pdf-signature-fixture.mjs";
const executar = promisify(execFile), nome = PDFName.of, reserva = 16384;
const hash = dados => createHash("sha256").update(dados).digest("hex");
async function opensslDisponivel() {
    for (const candidato of [process.env.OPENSSL_PATH, "C:/Program Files/Git/mingw64/bin/openssl.exe", "C:/Program Files/Git/usr/bin/openssl.exe", "openssl"].filter(Boolean)) {
        try { await executar(candidato, ["version"], { windowsHide: true, timeout: 10000 }); return candidato; } catch {}
    }
    throw new Error("OpenSSL necessário para os testes de assinaturas CMS reais; configure OPENSSL_PATH. Nenhuma dependência será instalada automaticamente.");
}
function novoValor(pdf, indice) {
    const placeholder = "/TEST_BYTE_RANGE_" + indice + "_" + "R".repeat(80);
    const valor = pdf.context.obj({ Type: "Sig", Filter: "Adobe.PPKLite", SubFilter: "adbe.pkcs7.detached", ByteRange: nome(placeholder.slice(1)),
        Contents: PDFHexString.of("0".repeat(reserva * 2)), Name: PDFString.of("4PrintShops TEST ONLY"), Reason: PDFString.of("Synthetic local test certificate; not a customer signature") });
    return { ref: pdf.context.register(valor), valor, placeholder };
}
function tamanhoDER(bytes) {
    if (bytes[0] !== 0x30) throw new Error("CMS sem SEQUENCE DER.");
    if (bytes[1] < 128) return 2 + bytes[1];
    const n = bytes[1] & 127;
    if (!n || n > 4) throw new Error("Comprimento DER não suportado na fixture.");
    let tamanho = 0; for (let i = 0; i < n; i++) tamanho = tamanho * 256 + bytes[2 + i];
    return 2 + n + tamanho;
}
async function preencherCMS(bytes, placeholder, contexto, id) {
    const texto = bytes.toString("latin1"), posicao = texto.indexOf(placeholder);
    if (posicao < 0 || texto.indexOf(placeholder, posicao + placeholder.length) >= 0) throw new Error("Placeholder ByteRange ambíguo.");
    const chave = texto.indexOf("/Contents <", posicao), inicio = texto.indexOf("<", chave), fim = texto.indexOf(">", inicio) + 1;
    if (chave < posicao || fim - inicio - 2 !== reserva * 2) throw new Error("Reserva Contents inesperada.");
    const faixa = [0, inicio, fim, bytes.length - fim], literal = `[ ${faixa.join(" ")} ]`;
    if (literal.length > placeholder.length) throw new Error("Reserva ByteRange insuficiente.");
    const pronto = Buffer.from(bytes); pronto.write(literal.padEnd(placeholder.length), posicao, "ascii");
    const dados = Buffer.concat([pronto.subarray(0, inicio), pronto.subarray(fim)]);
    const entrada = path.join(contexto.pasta, id + "-bytes.bin"), saida = path.join(contexto.pasta, id + "-cms.der");
    await writeFile(entrada, dados);
    await executar(contexto.openssl, ["cms", "-sign", "-binary", "-in", entrada, "-signer", contexto.certificado, "-inkey", contexto.chave, "-md", "sha256", "-outform", "DER", "-out", saida, "-nosmimecap"], { windowsHide: true, timeout: 30000 });
    const cms = await readFile(saida);
    if (cms.length > reserva) throw new Error("CMS excede a reserva da fixture.");
    pronto.write(cms.toString("hex").padEnd(reserva * 2, "0"), inicio + 1, "ascii");
    return pronto;
}
function acrescentarRevisao(bytes, pdf, campoRef, assinatura) {
    // Reescrever o PDF inteiro invalidaria a primeira assinatura. Aqui apenas
    // acrescentamos campo atualizado, novo Sig, xref parcial e trailer /Prev.
    const antigo = [...bytes.toString("latin1").matchAll(/startxref\s+(\d+)\s+%%EOF/g)].at(-1)?.[1];
    if (!antigo) throw new Error("xref anterior não localizado.");
    const campo = pdf.context.lookup(campoRef); campo.set(nome("V"), assinatura.ref);
    const prefixo = Buffer.from("\n", "ascii"), primeiroOffset = bytes.length + prefixo.length;
    const primeiro = Buffer.from(`${campoRef.objectNumber} ${campoRef.generationNumber} obj\n${campo}\nendobj\n`, "latin1");
    const segundoOffset = primeiroOffset + primeiro.length;
    const segundo = Buffer.from(`${assinatura.ref.objectNumber} 0 obj\n${assinatura.valor}\nendobj\n`, "latin1");
    const offsetXref = segundoOffset + segundo.length;
    const entradas = [[campoRef.objectNumber, campoRef.generationNumber, primeiroOffset], [assinatura.ref.objectNumber, 0, segundoOffset]].sort((a, b) => a[0] - b[0]);
    const linhas = entradas.map(([n, g, o]) => `${n} 1\n${String(o).padStart(10, "0")} ${String(g).padStart(5, "0")} n \n`).join("");
    const info = pdf.context.trailerInfo.Info ? `/Info ${pdf.context.trailerInfo.Info} ` : "";
    const trailer = Buffer.from(`xref\n${linhas}trailer\n<< /Size ${pdf.context.largestObjectNumber + 1} /Root ${pdf.context.trailerInfo.Root} ${info}/Prev ${antigo} >>\nstartxref\n${offsetXref}\n%%EOF\n`, "ascii");
    return Buffer.concat([bytes, prefixo, primeiro, segundo, trailer]);
}

export async function verificarAssinaturasCMS(bytes, contexto, id) {
    const pdf = await PDFDocument.load(bytes, { updateMetadata: false }), verificadas = [];
    for (const [ref, dict] of pdf.context.enumerateIndirectObjects()) {
        if (dict.lookupMaybe?.(nome("Type"), PDFName)?.asString() !== "/Sig") continue;
        const faixa = dict.lookup(nome("ByteRange")).asArray().map(v => v.asNumber()), conteudo = dict.lookup(nome("Contents")).asBytes();
        const cms = conteudo.slice(0, tamanhoDER(conteudo));
        const dados = Buffer.concat(Array.from({ length: faixa.length / 2 }, (_, i) => Buffer.from(bytes).subarray(faixa[i * 2], faixa[i * 2] + faixa[i * 2 + 1])));
        const base = path.join(contexto.pasta, `${id}-verificar-${ref.objectNumber}`);
        await writeFile(base + ".der", cms); await writeFile(base + ".bin", dados);
        await executar(contexto.openssl, ["cms", "-verify", "-binary", "-inform", "DER", "-in", base + ".der", "-content", base + ".bin", "-noverify", "-out", base + ".verificado"], { windowsHide: true, timeout: 30000 });
        verificadas.push({ objeto: ref.toString(), byteRange: faixa, revisaoAte: faixa.at(-2) + faixa.at(-1), sha256Conteudo: hash(dados), cmsVerificado: true });
    }
    return verificadas;
}

export async function criarConjuntoAssinaturasReais() {
    const pasta = await mkdtemp(path.join(tmpdir(), "4printshops-certificados-teste-"));
    const contexto = { pasta, openssl: await opensslDisponivel(), certificado: path.join(pasta, "certificado-teste.pem"), chave: path.join(pasta, "chave-efemera.pem") };
    await executar(contexto.openssl, ["req", "-x509", "-newkey", "rsa:2048", "-nodes", "-sha256", "-days", "1", "-subj", "/CN=4PrintShops TEST ONLY", "-addext", "basicConstraints=critical,CA:FALSE", "-addext", "keyUsage=critical,digitalSignature,nonRepudiation", "-keyout", contexto.chave, "-out", contexto.certificado], { windowsHide: true, timeout: 30000 }).catch(async erro => { await unlink(contexto.chave).catch(() => {}); throw erro; });
    const configuracoes = [
        { id: "real-visivel", total: 3, paginas: [1] },
        { id: "real-invisivel", total: 3, paginas: [2], invisivel: true },
        { id: "real-ultima", total: 8, paginas: [8] },
        { id: "real-multiplas", total: 4, paginas: [1, 4] },
        { id: "real-vazio", total: 3, paginas: [3], vazio: true },
        { id: "real-misto", total: 6, paginas: [2, 5] }
    ], casos = [];
    try {
        for (const config of configuracoes) {
            const pdf = await PDFDocument.load(await criarPDFAssinado({ ...config, vazia: true }), { updateMetadata: false });
            const campos = pdf.catalog.lookup(nome("AcroForm")).lookup(nome("Fields")).asArray();
            let bytes;
            if (config.vazio) bytes = Buffer.from(await pdf.save({ useObjectStreams: false, updateFieldAppearances: false }));
            else {
                const sig = novoValor(pdf, 1); pdf.context.lookup(campos[0]).set(nome("V"), sig.ref);
                bytes = await preencherCMS(Buffer.from(await pdf.save({ useObjectStreams: false, updateFieldAppearances: false })), sig.placeholder, contexto, config.id + "-1");
                for (let i = 1; i < campos.length; i++) {
                    const revisao = await PDFDocument.load(bytes, { updateMetadata: false }), proxima = novoValor(revisao, i + 1);
                    bytes = await preencherCMS(acrescentarRevisao(bytes, revisao, PDFRef.of(campos[i].objectNumber, campos[i].generationNumber), proxima), proxima.placeholder, contexto, config.id + "-" + (i + 1));
                }
            }
            const arquivo = path.join(pasta, config.id + ".pdf"); await writeFile(arquivo, bytes);
            const verificacoes = await verificarAssinaturasCMS(bytes, contexto, config.id);
            if (verificacoes.length !== (config.vazio ? 0 : config.paginas.length)) throw new Error("Contagem CMS inesperada na fixture.");
            casos.push({ ...config, arquivo, bytes, verificacoes, sha256: hash(bytes) });
        }
        await writeFile(path.join(pasta, "verificacoes.json"), JSON.stringify({ certificado: "Autoassinado RSA-2048/SHA-256, TEST ONLY, validade 1 dia, sem instalação no sistema", confiancaDaCadeiaVerificada: false, casos: casos.map(({ id, verificacoes, sha256 }) => ({ id, verificacoes, sha256 })) }, null, 2));
        return { ...contexto, casos };
    } finally { await unlink(contexto.chave).catch(() => {}); }
}
