export function renderSidebar(onToolSelected) {

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

    inicio.addEventListener("click", () => {

        onToolSelected("inicio");

    });


    //---------------------------------
    // QR Codes
    //---------------------------------

    const qrcodes = document.createElement("button");

    qrcodes.className = "menu-item";

    qrcodes.dataset.tool = "qrcodes";

    qrcodes.textContent = "QR Codes";

    menu.appendChild(qrcodes);

    qrcodes.addEventListener("click", () => {

        onToolSelected("qrcodes");

    });

    //---------------------------------
    // PDFs
    //---------------------------------

    const pdfs = document.createElement("button");

    pdfs.className = "menu-item";

    pdfs.dataset.tool = "pdfs";

    pdfs.textContent = "PDFs";

    menu.appendChild(pdfs);

    pdfs.addEventListener("click", () => {

        onToolSelected("pdfs");

    });


    sidebar.appendChild(menu);

}


//---------------------------------
// Destaca o item ativo no sidebar
//---------------------------------

export function setActiveTool(tool) {

    const itens = document.querySelectorAll(".menu-item");

    for (const item of itens) {

        item.classList.toggle(
            "active",
            item.dataset.tool === tool
        );

    }

}
