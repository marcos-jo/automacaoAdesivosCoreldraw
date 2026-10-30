import { medidaMm } from "./pdf-formats.js";

const formatador = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 2, useGrouping: false });
// Arredondamento apenas no indicador: os campos e o vínculo mantêm sua precisão.
export function formatarProporcao(largura, altura) {
    const fmt = n => formatador.format(n);
    const proporcao = largura >= altura ? `${fmt(largura / altura)}:1` : `1:${fmt(altura / largura)}`;
    return `${fmt(largura)} × ${fmt(altura)} mm · ≈ ${proporcao}`;
}

export class PDFProporcaoMedidas {
    constructor(buscar) {
        this.campo = id => buscar("#pdf-redim-" + id); this.ligada = false;
        this.clicar = () => {
            if (!this.ligada) try { this.medidas(); } catch { this.erro = "Informe largura e altura válidas antes de vincular."; this.exibir(); return; }
            this.ligada = !this.ligada; this.sincronizar();
        };
        this.campo("vincular").addEventListener("click", this.clicar); this.sincronizar();
    }
    medidas() { return [medidaMm(this.campo("largura").value), medidaMm(this.campo("altura").value)]; }
    sincronizar() {
        this.erro = null;
        try { const [a, b] = this.medidas(); this.razao = a / b; } catch { this.razao = null; }
        this.exibir();
    }
    alterar(lado) {
        this.erro = null;
        if (this.ligada && this.razao) try {
            const valor = medidaMm(this.campo(lado).value);
            const outro = lado === "largura" ? "altura" : "largura", nova = lado === "largura" ? valor / this.razao : valor * this.razao;
            medidaMm(String(nova)); this.campo(outro).value = String(Number(nova.toPrecision(15))).replace(".", ",");
        } catch { this.erro = "Informe uma medida válida; a medida vinculada deve ficar entre 0,1 e 5000 mm."; }
        if (!this.ligada) this.sincronizar(); else this.exibir();
    }
    validar() { if (this.erro) throw new Error(this.erro); }
    exibir() {
        let texto;
        try { texto = formatarProporcao(...this.medidas()); } catch { texto = "Proporção indisponível"; }
        this.campo("proporcao").textContent = this.erro || texto;
        const rotulo = this.ligada ? "Proporção vinculada" : "Proporção livre";
        this.campo("vincular").setAttribute("aria-label", rotulo);
        this.campo("vincular").title = rotulo + (this.ligada ? " — mantém a relação entre largura e altura da página final" : " — permite alterar largura e altura independentemente");
        this.campo("vincular").setAttribute("aria-pressed", String(this.ligada));
    }
    destruir() { this.campo("vincular").removeEventListener("click", this.clicar); }
}
