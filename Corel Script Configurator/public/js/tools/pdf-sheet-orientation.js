// Produz medidas efetivas; nunca modifica o cadastro que originou a folha.
export function orientarFolha(largura, altura, orientacao) {
    const a = Number(String(largura).trim().replace(',', '.'));
    const b = Number(String(altura).trim().replace(',', '.'));
    if (!Number.isFinite(a) || !Number.isFinite(b) || a <= 0 || b <= 0) return null;
    return orientacao === 'paisagem'
        ? { largura: Math.max(a, b), altura: Math.min(a, b) }
        : { largura: Math.min(a, b), altura: Math.max(a, b) };
}
