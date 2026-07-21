export class ScriptModel {

    constructor() {

        this.nome = "";
        this.caminho = "";
        this.versao = "";
        this.autor = "";

        this.sections = [];

    }

    addSection(section) {

        this.sections.push(section);

    }

}

export class Section {

    constructor(nome) {

        this.nome = nome;

        this.controls = [];

    }

    addControl(control) {

        this.controls.push(control);

    }

}

export class Control {

    constructor() {

        this.type = "";
        this.group = "";
        this.label = "";

        this.variable = "";
        this.value = null;

        this.unit = "";
        this.usaMM = false;

    }

    applyMetadata(metadata) {

        Object.assign(this, metadata);

    }

}