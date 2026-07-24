import { readFileSync, existsSync } from "node:fs";
import path from "node:path";


import {
    listarScripts,
    carregarScript
} from "./repository.js";

import http from "node:http";

export function startServer() {

    const server = http.createServer(handleRequest);

    server.listen(3000, () => {

        console.log("==================================");
        console.log(" Corel Script Configurator");
        console.log("==================================");
        console.log("");
        console.log("Servidor iniciado.");
        console.log("http://localhost:3000");

    });

}

function handleRequest(req, res) {

    console.log(
    `[${new Date().toLocaleTimeString()}] ${req.method} ${req.url}`
    );

    if (req.url === "/") {

    serveIndex(res);

    return;

    }

    if (req.url === "/api/scripts") {

        const scripts = listarScripts();

        res.writeHead(200, {
            "Content-Type": "application/json; charset=utf-8"
        });

        res.end(JSON.stringify({
            scripts
        }));

        return;

    }

    if (req.url.startsWith("/api/scripts/")) {

    try {

        const nomeArquivo = decodeURIComponent(
            req.url.substring("/api/scripts/".length)
        );

        const model = carregarScript(nomeArquivo);

        res.writeHead(200, {
            "Content-Type": "application/json; charset=utf-8"
        });

        res.end(JSON.stringify(model, null, 4));

    }

    catch (erro) {

        res.writeHead(404, {
            "Content-Type": "application/json"
        });

        res.end(JSON.stringify({
            erro: erro.message
        }));

    }

    return;

}

    if (serveStatic(req, res))
    return;

    res.writeHead(404);

    res.end("Rota não encontrada");

}

function serveIndex(res) {

    const arquivo = "./public/index.html";

    if (!existsSync(arquivo)) {

        res.writeHead(404);

        res.end("index.html não encontrado.");

        return;

    }

    const html = readFileSync(arquivo, "utf8");

    res.writeHead(200, {

        "Content-Type": "text/html; charset=utf-8"

    });

    res.end(html);

}

function serveStatic(req, res) {

    const arquivo = path.join("public", req.url);

    if (!existsSync(arquivo))
        return false;

    const extensao = path.extname(arquivo);

    const tipos = {

        ".css": "text/css",
        ".js": "text/javascript",
        ".html": "text/html",

        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".svg": "image/svg+xml",
        ".ico": "image/x-icon"

    };

    const mime = tipos[extensao] || "application/octet-stream";

    const arquivoBuffer = readFileSync(arquivo);

    res.writeHead(200, {

        "Content-Type": mime

    });

    res.end(arquivoBuffer);

    return true;

}