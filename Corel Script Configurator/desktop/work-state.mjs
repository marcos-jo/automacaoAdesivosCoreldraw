import {randomUUID} from 'node:crypto';

// Falha fechada se a janela recarregar, não responder ou entregar dados inválidos.
export function workStateRequest(getWindow,{timeout=5000}={}) {
    const pending=new Map();
    const receive=(event,token,state)=>{
        const task=pending.get(token);
        if(!task||event.sender!==task.contents||event.senderFrame!==task.contents.mainFrame)return;
        pending.delete(token);clearTimeout(task.timer);
        task.resolve(state&&typeof state.pending==='boolean'&&typeof state.busy==='boolean'?state:{pending:true,busy:true});
    };
    const request=()=>new Promise(resolve=>{
        const window=getWindow();if(!window||window.isDestroyed()){resolve({pending:true,busy:true});return;}
        const token=randomUUID(),contents=window.webContents;
        const timer=setTimeout(()=>{pending.delete(token);resolve({pending:true,busy:true});},timeout);
        pending.set(token,{contents,resolve,timer});
        try{contents.send('updates:work-request',token);}catch{clearTimeout(timer);pending.delete(token);resolve({pending:true,busy:true});}
    });
    const destroy=()=>{for(const task of pending.values()){clearTimeout(task.timer);task.resolve({pending:true,busy:true});}pending.clear();};
    return {request,receive,destroy};
}
