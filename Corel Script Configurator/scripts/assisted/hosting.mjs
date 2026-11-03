import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {writeFile} from 'node:fs/promises';
import {auditRepository} from './publish.mjs';

// Laboratório somente loopback; não instala serviço, certificado ou regra de firewall.
export async function labCaddyfile(root,destination){
    await auditRepository(root);
    const quoted=JSON.stringify(path.resolve(root).replaceAll('\\','/'));
    const config=`{
    admin off
    auto_https off
    persist_config off
}
http://127.0.0.1:18084 {
    bind 127.0.0.1
    root * ${quoted}
    header Cache-Control "no-store"
    header X-Content-Type-Options "nosniff"
    @public {
        method GET HEAD
        path_regexp publication ^/(channels/internal(-lab)?[.]json|releases/[0-9A-Za-z.+-]+/[0-9A-Za-z_-][0-9A-Za-z._-]*[.]exe)$
    }
    handle @public {
        file_server
    }
    handle {
        respond "Não encontrado" 404
    }
}
`;
    await writeFile(destination,config);return destination;
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
    if(!process.argv[3])throw new Error('Uso: node scripts/assisted/hosting.mjs <repositorio> <Caddyfile-lab>');
    console.log(await labCaddyfile(process.argv[2],process.argv[3]));
}
