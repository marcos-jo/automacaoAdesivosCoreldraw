import { PDFDocument, PDFName, PDFSignature } from "/vendor/pdf-lib/pdf-lib.js";
import { importarImagem, LIMITE_COMBINACAO } from "./pdf-image-import.js";

// Rascunho independente. O workspace só recebe o lote após Combinar.
export class PDFCombinacao {
    constructor() {
        this.sources = [];
        this.pages = [];
        this.serial = 0;
        this.size = 0;
    }

    async add(files) {
        const total = files.reduce((n, file) => n + (file.size || 0), this.size);
        if (total > LIMITE_COMBINACAO) throw new Error("A seleção ultrapassa o limite de 512 MiB.");
        const sources = [], pages = [];
        for (const file of files) {
            try {
                const extension = file.name.split(".").at(-1).toLowerCase();
                if (!["pdf", "jpg", "jpeg", "png"].includes(extension)) throw new Error("Tipo incompatível. Selecione PDF, JPG/JPEG ou PNG.");
                const bytes = new Uint8Array(await file.arrayBuffer());
                if (this.size + sources.reduce((n, source) => n + source.size, 0) + bytes.length > LIMITE_COMBINACAO) throw new Error("A seleção ultrapassa 512 MiB.");
                const pdf = extension === "pdf"
                    ? await PDFDocument.load(bytes, { updateMetadata: false })
                    : (await importarImagem(bytes, extension)).pdf;
                if (!pdf.getPageCount()) throw new Error("Documento sem páginas.");
                if (pdf.catalog.has(PDFName.of("OCProperties"))) throw new Error("PDF com camadas opcionais. A aparência da combinação ainda não foi validada para esse recurso.");
                if (pdf.catalog.has(PDFName.of("OutputIntents"))) throw new Error("PDF com intenção de saída/perfil de impressão. A combinação ainda não preserva esse perfil global com segurança.");
                const form = pdf.catalog.get(PDFName.of("AcroForm"));
                if (form && pdf.getForm().getFields().some(field => !(field instanceof PDFSignature))) {
                    throw new Error("Formulários exigem uma cópia para impressão conferida antes da combinação.");
                }
                for (const page of pdf.getPages()) for (const ref of page.node.Annots()?.asArray() || []) {
                    let widget = pdf.context.lookup(ref);
                    if (widget.get(PDFName.of("Subtype"))?.toString() !== "/Widget") continue;
                    const seen = new Set();
                    while (!widget.has(PDFName.of("FT")) && widget.has(PDFName.of("Parent")) && !seen.has(widget)) {
                        seen.add(widget); widget = pdf.context.lookup(widget.get(PDFName.of("Parent")));
                    }
                    if (widget.get(PDFName.of("FT"))?.toString() !== "/Sig") throw new Error("Campo de formulário sem suporte seguro à combinação. Prepare uma cópia para impressão conferida.");
                }
                const previewBytes = extension === "pdf" ? bytes : await pdf.save();
                const source = { id: ++this.serial, name: file.name, type: extension.toUpperCase(), pdf, previewBytes, size: bytes.length, total: pdf.getPageCount() };
                sources.push(source);
                for (let index = 0; index < source.total; index++) {
                    const page = pdf.getPage(index), box = page.getCropBox();
                    const rotated = Math.abs(page.getRotation().angle % 180) === 90;
                    const unit = page.node.get(PDFName.of("UserUnit"))?.asNumber() || 1;
                    pages.push({ id: ++this.serial, source, index, width: (rotated ? box.height : box.width) * unit * 25.4 / 72, height: (rotated ? box.width : box.height) * unit * 25.4 / 72 });
                }
            } catch (error) {
                throw new Error(`${file.name} — ${error.message}`);
            }
        }
        this.sources.push(...sources);
        this.pages.push(...pages);
        this.size += sources.reduce((n, source) => n + source.size, 0);
    }

    move(from, to) {
        if (![from, to].every(n => Number.isInteger(n) && n >= 0 && n < this.pages.length)) throw new Error("Posição inválida.");
        if (from === to) return;
        const [page] = this.pages.splice(from, 1);
        this.pages.splice(to, 0, page);
    }

    remove(id, wholeFile = false) {
        const page = this.pages.find(p => p.id === id);
        if (!page) return;
        this.pages = this.pages.filter(p => wholeFile ? p.source !== page.source : p !== page);
        this.sources = this.sources.filter(source => this.pages.some(p => p.source === source));
        this.size = this.sources.reduce((n, source) => n + source.size, 0);
    }

    async build() {
        if (!this.pages.length) throw new Error("Adicione pelo menos uma página.");
        const pdf = await PDFDocument.create();
        // Cada origem é copiada uma vez, mantendo recursos compartilhados entre suas páginas.
        const copied = new Map();
        for (const source of this.sources) {
            const indices = this.pages.filter(p => p.source === source).map(p => p.index);
            const pages = await pdf.copyPages(source.pdf, indices);
            copied.set(source, new Map(indices.map((index, i) => [index, pages[i]])));
        }
        for (const page of this.pages) pdf.addPage(copied.get(page.source).get(page.index));
        const bytes = await pdf.save();
        if (bytes.length > LIMITE_COMBINACAO) throw new Error("O documento combinado ultrapassa 512 MiB.");
        return bytes;
    }

    origins() {
        return this.sources.map(source => ({ pdf: source.pdf, nome: source.name, tamanho: source.size,
            indices: Array.from({ length: source.total }, (_, index) => this.pages.findIndex(page => page.source === source && page.index === index)) }));
    }
}
