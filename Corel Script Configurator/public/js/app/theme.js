const KEY = '4printshops.theme';
export function aplicarTema(value, storage, root = document.documentElement) {
    const theme = value === 'dark' ? 'dark' : 'light';
    root.dataset.theme = theme;
    try { (storage ?? globalThis.localStorage)?.setItem(KEY, theme); } catch {}
    return theme;
}
export function iniciarTema(storage, root = document.documentElement) {
    let value;
    try { value = (storage ?? globalThis.localStorage)?.getItem(KEY); } catch {}
    return aplicarTema(value, storage, root);
}
export function menuTema(container) {
    const menu = document.createElement('details');
    menu.id = 'app-theme-menu';
    menu.innerHTML = '<summary>Tema</summary><div role="group" aria-label="Tema"><button type="button" data-theme-choice="light">Claro</button><button type="button" data-theme-choice="dark">Escuro</button></div>';
    const update = () => { for (const button of menu.querySelectorAll('button')) button.setAttribute('aria-pressed', String(button.dataset.themeChoice === document.documentElement.dataset.theme)); };
    menu.addEventListener('click', event => {
        const button = event.target.closest('button[data-theme-choice]');
        if (!button) return;
        aplicarTema(button.dataset.themeChoice); update();
    });
    update(); container.append(menu);
}
