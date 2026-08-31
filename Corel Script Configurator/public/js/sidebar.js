export function renderSidebar() {

    const sidebar = document.getElementById("sidebar");

    sidebar.innerHTML = "";

    //---------------------------------
    // Título
    //---------------------------------

    const titulo = document.createElement("div");

    titulo.className = "sidebar-title";

    titulo.textContent = "Ferramentas";

    sidebar.appendChild(titulo);


    //---------------------------------
    // Menu
    //---------------------------------

    const menu = document.createElement("nav");

    menu.className = "sidebar-menu";


    //---------------------------------
    // Início
    //---------------------------------

    const inicio = document.createElement("button");

    inicio.className = "menu-item active";

    inicio.dataset.tool = "inicio";

    inicio.textContent = "Início";

    menu.appendChild(inicio);


    //---------------------------------
    // Adesivos
    //---------------------------------

    const adesivos = document.createElement("button");

    adesivos.className = "menu-item";

    adesivos.dataset.tool = "adesivos";

    adesivos.textContent = "Adesivos";

    menu.appendChild(adesivos);


    //---------------------------------
    // QR Codes
    //---------------------------------

    const qrcodes = document.createElement("button");

    qrcodes.className = "menu-item";

    qrcodes.dataset.tool = "qrcodes";

    qrcodes.textContent = "QR Codes";

    menu.appendChild(qrcodes);


    sidebar.appendChild(menu);

}