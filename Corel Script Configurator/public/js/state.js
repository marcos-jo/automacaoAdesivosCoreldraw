class ApplicationState {

    constructor() {

        this.script = null;

    }

    setScript(model) {

        this.script = model;

    }

    getScript() {

        return this.script;

    }

}

export const state = new ApplicationState();