import {configuredUpdater} from './updates.mjs';

// Uma consulta em voo; inicialização silenciosa, download e instalação explícitos.
export function updateAction(app,{getWindow,dialog,canInstall,publish=()=>{},lock=()=>{},configure=configuredUpdater}) {
    let controller,flight=false,disposed=false,state={phase:'idle'};
    const emit=value=>{state=value;if(!disposed)publish(value);};
    const message=text=>dialog.showMessageBox(getWindow(),{type:'info',title:'Atualizações do 4PrintShops',message:text,buttons:['Continuar']});
    const confirm=async()=>{
        const result=await dialog.showMessageBox(getWindow(),{type:'question',title:'Instalar atualização e reiniciar',
            message:'Instalar a atualização e reiniciar o 4PrintShops agora?',detail:'As preferências serão mantidas. Documentos abertos não são recuperados após reiniciar.',
            buttons:['Instalar e reiniciar','Mais tarde'],defaultId:1,cancelId:1,noLink:true});
        return result.response===0;
    };
    const observe=value=>{
        if(value.phase==='error')lock(false);
        emit(value);const w=getWindow();if(w&&!w.isDestroyed())w.setProgressBar(value.phase==='checking'?2:value.phase==='downloading'?(value.percent||0)/100:-1);
    };
    async function ensure(){
        if(!controller){controller=await configure(app,{canInstall,confirmInstall:confirm});controller?.on('state',observe);}
        return controller;
    }
    const run=async({silent=false}={})=>{
        if(flight||disposed)return state;flight=true;
        try{
            if(!await ensure()){emit({phase:'disabled'});if(!silent)await message('Atualizações não configuradas nesta versão. Nenhuma consulta à internet foi realizada.');return state;}
            const value=await controller.check();
            if(disposed||silent)return state;
            if(value.phase==='current'){await message('Nenhuma atualização disponível.');return state;}
            if(value.phase==='error'){await message('Não foi possível concluir a atualização: '+value.message+' Tente novamente pelo menu.');return state;}
            if(value.phase==='available'){
                const answer=await dialog.showMessageBox(getWindow(),{type:'question',title:'Nova versão encontrada',message:`4PrintShops ${value.version}`,
                    detail:value.notes||'Novidades e correções não informadas.',buttons:['Atualizar agora','Mais tarde'],defaultId:1,cancelId:1,noLink:true});
                if(disposed||answer.response!==0){controller.later();return state;}
                await controller.download();
                if(controller.state.phase==='error'){await message('Download não concluído: '+controller.state.message+' Tente novamente pelo menu.');return state;}
            }
            if(!disposed&&controller.state.phase==='ready'){
                lock(true);
                try{await controller.install();}finally{if(controller.state.phase!=='installing')lock(false);}
            }
        }catch(error){lock(false);if(!disposed){emit({phase:'error',message:error.message,code:error.code||''});if(!silent)await message('Não foi possível concluir a atualização: '+error.message);}}
        finally{flight=false;}
        return state;
    };
    run.start=()=>run({silent:true});
    run.getState=()=>state;
    run.destroy=()=>{disposed=true;controller?.off('state',observe);lock(false);};
    return run;
}
