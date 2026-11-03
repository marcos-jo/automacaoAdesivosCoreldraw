import http from 'node:http';
import path from 'node:path';
import {stat} from 'node:fs/promises';
import {createReadStream} from 'node:fs';
import {Transform} from 'node:stream';
import {fileURLToPath} from 'node:url';

// Somente loopback e artefatos gerados; não publica arquivos do projeto/usuário.
export function serveUpdateLab(root,{port=18083,onRequest=()=>{},bytesPerSecond=0}={}){
    const server=http.createServer(async(req,res)=>{
        try{
            const name=decodeURIComponent(new URL(req.url,'http://127.0.0.1').pathname).slice(1);
            if(!['GET','HEAD'].includes(req.method)||!/^experimental\.yml$|^4PrintShops-Update-Lab-1\.0\.0-experimental\.[23]-Setup\.exe(?:\.blockmap)?$/.test(name)){res.writeHead(404);res.end();return;}
            const file=path.join(root,name),info=await stat(file);
            let start=0,end=info.size-1,status=200;
            if(req.headers.range){const range=/^bytes=(\d+)-(\d*)$/.exec(req.headers.range);if(!range){res.writeHead(416);res.end();return;}start=Number(range[1]);end=range[2]?Number(range[2]):end;status=206;
                if(start>end||end>=info.size){res.writeHead(416);res.end();return;}}
            const headers={'Content-Length':end-start+1,'Cache-Control':'no-store','Accept-Ranges':'bytes','Content-Type':name.endsWith('.yml')?'text/yaml':'application/octet-stream'};
            if(status===206)headers['Content-Range']=`bytes ${start}-${end}/${info.size}`;
            onRequest({name,status,start,end});res.writeHead(status,headers);
            if(req.method==='HEAD')res.end();else{
                const stream=createReadStream(file,{start,end});stream.on('error',()=>res.destroy());res.on('close',()=>stream.destroy());
                if(bytesPerSecond>0&&name.endsWith('.exe')){
                    const slow=new Transform({transform(chunk,_encoding,done){setTimeout(()=>done(null,chunk),1000*chunk.length/bytesPerSecond);}});
                    slow.on('error',()=>res.destroy());res.on('close',()=>slow.destroy());stream.pipe(slow).pipe(res);
                }else stream.pipe(res);
            }
        }catch{res.writeHead(404);res.end();}
    });
    return new Promise((resolve,reject)=>{server.once('error',reject);server.listen(port,'127.0.0.1',()=>resolve(server));});
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
    const root=path.resolve(process.argv[2]||'dist-desktop/update-lab/1.0.0-experimental.3');
    const server=await serveUpdateLab(root);console.log('Laboratório: http://127.0.0.1:18083/ — '+root);
    for(const signal of ['SIGINT','SIGTERM'])process.on(signal,()=>server.close());
}
