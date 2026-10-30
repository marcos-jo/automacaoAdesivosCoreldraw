import {configuredUpdater} from './updates.mjs';

export function updateAction(app,{getWindow,dialog,canInstall}) {
    let controller,flight=false;
    const message=(text)=>dialog.showMessageBox(getWindow(),{type:'info',title:'Atualizações do 4PrintShops',message:text,buttons:['Continuar']});
    const confirm=async()=>{
        const result=await dialog.showMessageBox(getWindow(),{type:'question',title:'Atualização pronta',message:'A cópia atual será encerrada para instalar a atualização.',buttons:['Atualizar agora','Mais tarde'],defaultId:1,cancelId:1});
        return result.response===0;
    };
    return async()=>{
        if(flight)return;flight=true;
        try {
            controller ||= await configuredUpdater(app,{canInstall,confirmInstall:confirm});
            if(!controller){await message('Atualizações não configuradas nesta versão experimental. Nenhuma consulta à internet foi realizada. A distribuição futura requer provedor e identidade de assinatura configurados.');return;}
            if(!controller.observed){controller.observed=true;controller.on('state',state=>{
                const w=getWindow();if(!w||w.isDestroyed())return;
                w.setProgressBar(state.phase==='checking'?2:state.phase==='downloading'?(state.percent||0)/100:-1);
            });}
            let state=controller.state.phase==='ready'?controller.state:await controller.check();
            if(state.phase==='current'){await message('Nenhuma atualização disponível.');return;}
            if(state.phase==='available'){
                const answer=await dialog.showMessageBox(getWindow(),{type:'question',title:'Nova versão encontrada',message:`4PrintShops ${state.version}`,detail:state.notes||'Novidades e correções não informadas.',buttons:['Atualizar agora','Mais tarde'],defaultId:1,cancelId:1});
                if(answer.response!==0){controller.later();return;}
                state=await controller.download();
            }
            if(state.phase==='error')throw new Error(state.message);
            if(state.phase==='ready')await controller.install();
        }catch(error){await message('Não foi possível concluir a atualização: '+error.message+' Tente novamente pela opção Verificar atualizações.');}
        finally{flight=false;const window=getWindow();if(window&&!window.isDestroyed())window.setProgressBar(-1);}
    };
}
