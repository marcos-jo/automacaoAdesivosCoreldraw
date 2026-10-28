// Atualiza os mesmos elementos técnicos sem reconstruir toda a lista a cada giro.
export class PDFInformacoesDocumento {
    constructor(buscar) {
        this.buscar = buscar;
        this.textos = [];
        this.paginas = [];
    }

    atualizar(paginas, nome) {
        const buscar = this.buscar;
        const primeira = paginas[0];
        const variados = paginas.some(p => Math.abs(p.largura - primeira.largura) > .1 || Math.abs(p.altura - primeira.altura) > .1);
        const formatar = n => n.toLocaleString("pt-BR", { maximumFractionDigits: 2 });
        buscar("#pdf-nome").textContent = nome;
        buscar("#pdf-conteudo").hidden = false;
        buscar("#pdf-total").textContent = paginas.length;
        buscar("#pdf-tamanho").textContent = variados ? "Variados" : `${formatar(primeira.largura)} × ${formatar(primeira.altura)} mm`;
        const informacoes = buscar("#pdf-informacoes");
        informacoes.classList.toggle("pdf-alerta", variados);
        informacoes.textContent = variados ? "⚠ Informações · 1 alerta" : "Informações";
        buscar("#pdf-alertas").textContent = variados ? "Atenção: páginas com dimensões diferentes." : "Nenhuma diferença de tamanho encontrada.";
        const lista = buscar("#pdf-lista-paginas");
        while (lista.children.length > paginas.length) lista.children[lista.children.length - 1].remove();
        const textos = paginas.map((p, indice) => {
            const anterior = this.paginas[indice];
            return anterior?.largura === p.largura && anterior.altura === p.altura
                ? this.textos[indice]
                : `Página ${p.numero}: ${formatar(p.largura)} × ${formatar(p.altura)} mm`;
        });
        textos.forEach((texto, indice) => {
            if (!lista.children[indice]) lista.appendChild(document.createElement("li"));
            if (texto !== this.textos[indice]) lista.children[indice].textContent = texto;
        });
        this.textos = textos;
        this.paginas = paginas;
    }
}
