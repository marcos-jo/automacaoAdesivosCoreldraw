// Apenas opções da janela; nunca armazena ou altera PDF/DEVMODE.
export const camposPorModo=['paginas','nupSelecao','nupDistribuicao','grade','colunas','linhas','ordem','quantidade','rotacao','rotacaoAutomatica','escalaModo','escala','pecaLargura','pecaAltura','esquerda','direita','superior','inferior','espacoX','espacoY','sangria','comprimento','afastamento','espessura','modoMarcas','centralizar'];
export class ConfiguracoesMontagem {
    constructor(values){this.modo='repeat';this.estados={repeat:{values:{...values},fator:1,base:null},nup:{values:{...values,nupSelecao:'todas',grade:'4',paginas:'1',escalaModo:'real',pecaLargura:'',pecaAltura:''},fator:1,base:null}};}
    trocar(modo,values,fator=1,base=null){
        if(!Object.hasOwn(this.estados,modo))throw Error('Tipo de montagem inválido.');
        this.estados[this.modo]={values:{...values},fator,base:base?{...base}:null};this.modo=modo;
        const saved=this.estados[modo];return {values:{...saved.values},fator:saved.fator,base:saved.base?{...saved.base}:null};
    }
    limparAjustes(){for(const state of Object.values(this.estados)){if(state.base)Object.assign(state.values,state.base);state.fator=1;state.base=null;}}
}
