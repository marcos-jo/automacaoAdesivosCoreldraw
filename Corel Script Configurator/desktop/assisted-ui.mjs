import {configuredAssisted} from './assisted-updates.mjs';

// Separado do electron-updater: somente revela o arquivo, nunca executa ou reinicia.
export function assistedAction(app,pkg,{getWindow,dialog,shell,canPrepare,lock=()=>{},publish=()=>{},configure=configuredAssisted}){
    let controller,flight=false,disposed=false,state={phase:'idle'};
    const emit=value=>{state=value;if(!disposed)publish(value);};
    const message=text=>dialog.showMessageBox(getWindow(),{title:'Atualização assistida pela rede',type:'info',message:text,buttons:['Continuar']});
    async function ensure(){if(!controller){controller=await configure(app,pkg);controller?.on('state',emit);}return controller;}
    async function run(action='check',{silent=false}={}){
        if(flight||disposed)return state;
        if(!['check','details','download','prepare'].includes(action))throw new Error('Ação de atualização assistida inválida.');
        flight=true;
        try{
            if(!await ensure()){emit({phase:'disabled'});if(!silent)await message('Distribuição interna assistida não configurada nesta instalação.');return state;}
            if(action==='check')await controller.check();
            if(silent||disposed)return state;
            if(action==='download'){await controller.download();return state;}
            if(action==='prepare'){
                // Duas consultas frescas e trava compartilhada com importação/exportação.
                lock(true);
                try{
                    if(!await canPrepare()){await message('Exporte as alterações de todas as abas e conclua as operações antes de preparar a instalação. Nenhum documento foi fechado.');return state;}
                    await controller.preparedFile();
                    const answer=await dialog.showMessageBox(getWindow(),{title:'Preparar instalação assistida',type:'question',
                        message:`Localizar o instalador verificado da versão ${controller.state.version}?`,
                        detail:'A assinatura da publicação e a integridade do arquivo foram verificadas. Isso não substitui a assinatura Authenticode do Windows. Salve seu trabalho, feche o 4PrintShops e execute o instalador conscientemente pelo Explorador. O aplicativo não executará o arquivo nem reiniciará.',
                        buttons:['Localizar instalador','Mais tarde'],defaultId:1,cancelId:1,noLink:true});
                    if(disposed||answer.response!==0)return state;
                    if(!await canPrepare()){await message('Há alterações ou operações pendentes. A instalação foi adiada.');return state;}
                    const file=await controller.preparedFile();
                    if(!disposed)shell.showItemInFolder(file);
                }finally{lock(false);}
                return state;
            }
            if(state.phase==='available'){
                const answer=await dialog.showMessageBox(getWindow(),{title:'Nova versão na rede interna',type:'question',message:`4PrintShops ${state.version} — atualização assistida`,
                    detail:[...state.notes,'','Publicação autenticada. O download não instala nem fecha o aplicativo.'].join('\n'),
                    buttons:['Baixar instalador','Mais tarde'],defaultId:1,cancelId:1,noLink:true});
                if(!disposed&&answer.response===0)await controller.download();
            }else if(state.phase==='ready')await message(`Versão ${state.version} verificada. Use “Preparar instalação” no aviso para localizar o arquivo. A instalação é manual e requer fechar o aplicativo após exportar os PDFs.`);
            else if(state.phase==='error')await message('Atualização indisponível: '+state.message+' Nenhum instalador será oferecido.');
            else if(state.phase==='current')await message('Nenhuma versão mais nova disponível na rede interna.');
        }catch(error){if(!disposed){emit({phase:'error',message:error.message});if(!silent)await message('Não foi possível preparar a atualização: '+error.message);}}
        finally{flight=false;}
        return state;
    }
    run.start=()=>run('check',{silent:true});
    run.getState=()=>state;
    run.destroy=()=>{disposed=true;controller?.off('state',emit);};
    return run;
}
