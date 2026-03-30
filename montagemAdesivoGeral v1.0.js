// =======================================================
// SCRIPT ADESIVO REDONDO tamanho personalizável
// CorelDRAW 2021
// =======================================================

function main() {

    var doc = host.ActiveDocument;
    var page = doc.ActivePage;
    
    doc.ClearSelection(); // Limpa seleção para evitar conflito com o resto do script
    
    function mm(valor) {
        return valor / 25.4;
    } // Converte medida interna do corel para milímetros
    
    
    //-----------------------------
    // CONFIGURAÇÕES
    var corteRedondo = true; // true para sim e false para não.
    
    var corteQuadrado = false; // true para sim e false para não.

    var cortePersonalizado = false; // true para sim e false para não, caso queira usar o corte personalizado, é necessário criar deixar o corte personalizado já na camada de corte e marcar as opções coreRedondo e corteQuadrado como false.
    
    var etiquetaEscolar = false; // true para sim e false nao. 

    var tamanhoHorizontal = mm(20); // largura do adesivo.
    
    var tamanhoVertical = mm(20); // altura do adesivo.
    // -----------------------------


    // -----------------------------
    // Caso utilize o corte personalizado, preencha esses valores abaixo
    
    var quantidadeDeCopiasHorizontal = 10; // Quantidade de copias para a direita
    
    var quantidadeDeCopiasVertical = 2; // Quantidade de cópias para cima
    
    var posicaoInicialX = mm(100); // Posição inicial X do primeiro adesivo
    
    var posicaoInicialY = mm(35); // Posição inicial Y do primeiro adesivo
    
    //-----------------------------

    //-----------------------------
    // Compensando o posicionamento dos itens diferente nas medidas internas do Corel, para que o posicionamento fique correto
    // a solução que encontrei foi descobrir o raio e adicionar ele no eixo x e encontrar o raio e subtrair no eixo y.
    // Não pergunta como eu descobri e como achei a solução, só confia que funciona. 😀
    var compensandoRaioHorizontal = tamanhoHorizontal / 2;
    var compensandoRaioVertical = tamanhoVertical / 2;
    posicaoInicialX = posicaoInicialX - compensandoRaioHorizontal;
    posicaoInicialY = posicaoInicialY + compensandoRaioVertical;
    //-----------------------------
 
    //-----------------------------
    // criação do scrypt que distribui os adesivos nas quantidades necessárias para preencher a área útil da folha, tentando esquivar do QRcode
    posicaoInicialX = mm(5) + tamanhoHorizontal;
    posicaoinicialY = mm(5) + tamanhoVertical;

    //Duplicatas até o limite dá área útil da folha
    var duplicatosX = 0;
    var duplicatosY = 0;
    var contagemX = tamanhoHorizontal + mm(1);
    var contagemY = tamanhoVertical + mm(1);

    while (contagemX <= (mm(318) - (tamanhoHorizontal + mm(1)))) {
        duplicatosX++;
        contagemX = contagemX + tamanhoHorizontal;
    }

    quantidadeDeCopiasHorizontal = duplicatosX;

    //-----------------------------

    if (corteRedondo) {
        
        /*

        if (corteRedondo && tamanhoHorizontal == mm(20) && tamanhoVertical == mm(20)) {
            posicaoInicialX = mm(18);
            posicaoInicialY = mm(15);
            quantidadeDeCopiasHorizontal = 14;
            quantidadeDeCopiasVertical = 20;
        }

        if (corteRedondo && tamanhoHorizontal == mm(25) && tamanhoVertical == mm(25)) {
            posicaoInicialX = mm(22);
            posicaoInicialY = mm(32);
            quantidadeDeCopiasHorizontal = 14;
            quantidadeDeCopiasVertical = 20;
        }

        if (corteRedondo && tamanhoHorizontal == mm(30) && tamanhoVertical == mm(30)) {
            posicaoInicialX = mm(25.5);
            posicaoInicialY = mm(38.5);
            quantidadeDeCopiasHorizontal = 14;
            quantidadeDeCopiasVertical = 20;
        }

        if (corteRedondo && tamanhoHorizontal == mm(35) && tamanhoVertical == mm(35)) {
            posicaoInicialX = mm(35.5);
            posicaoInicialY = mm(36.5);
            quantidadeDeCopiasHorizontal = 14;
            quantidadeDeCopiasVertical = 20;
        }

        if (corteRedondo && tamanhoHorizontal == mm(40) && tamanhoVertical == mm(40)) {
            posicaoInicialX = mm(42);
            posicaoInicialY = mm(35);
            quantidadeDeCopiasHorizontal = 14;
            quantidadeDeCopiasVertical = 20;
        }

        if (corteRedondo && tamanhoHorizontal == mm(45) && tamanhoVertical == mm(45)) {
            posicaoInicialX = mm(50);
            posicaoInicialY = mm(33);
            quantidadeDeCopiasHorizontal = 14;
            quantidadeDeCopiasVertical = 20;
        }

        if (corteRedondo && tamanhoHorizontal == mm(50) && tamanhoVertical == mm(50)) {
            posicaoInicialX = mm(37.5);
            posicaoInicialY = mm(36);
            quantidadeDeCopiasHorizontal = 14;
            quantidadeDeCopiasVertical = 20;
        }

        */
        
        var deslocamentoHorizontal = tamanhoHorizontal + mm(1); // 1mm de espaço entre os adesivos
        var deslocamentoVertical = tamanhoVertical + mm(1); // 1mm de espaço entre os adesivos

        //-----------------------------
        // Agrupa os itens da Camada 1 caso esteja com arquivos diversos, isso evita bugs

        var layer = page.Layers.Item("Camada 1"); // Agrupa todos os itens da camada 1 para facilitar a manipulação     
        var range = host.CreateShapeRange(); // Cria um ShapeRange vazio
        
        for (var i = 1; i <= layer.Shapes.Count; i++) {
            range.Add(layer.Shapes.Item(i));
        } // Adiciona todos os shapes da camada ao range
        
        var grupo = range.Group(); // Agrupa tudo

        //grupo.LockAspectRatio = true; // trava a proporção

        grupo.SizeWidth = deslocamentoHorizontal; // Ajuste de tamanho Horizontal X com a sangria de 1mm da var deslocamentoHorizontal
        grupo.SizeHeight = deslocamentoVertical; // Ajuste de tamanho Vertical Y com a sangria de 1mm da var deslocamentoVertical

        doc.ClearSelection(); // Limpa seleção para não conflitar com o resto do script

        //-----------------------------
        //Cria o círculo na camada de corte
        
        var layer = host.ActiveDocument.ActivePage.Layers.Item("Cut Layer");
        var circulo = layer.CreateEllipse2(
        mm(150),   // centro X
        mm(150),   // centro Y
        compensandoRaioHorizontal,  // raio horizontal
        compensandoRaioVertical   // raio vertical
        );
        // diâmetro = 20mm → raio = 10mm
        // -----------------------------

        range = host.CreateShapeRange(); // Esvazia o range para reutilizar

        // -----------------------------
        // OBJETOS BASE
        // -----------------------------
        var imgBase   = page.Layers.Item("Camada 1").Shapes.Item(1);
        var corteBase = page.Layers.Item("Cut Layer").Shapes.Item(1);

        // -----------------------------
        // POSIÇÃO INICIAL
        // -----------------------------

        // Define o ponto de referência como o centro (cdrCenter = 5)
        // No JS do Corel, dá para usar o valor numérico se a constante não estiver mapeada
        host.ActiveDocument.ReferencePoint = 5; 
        //imgBase.SetPosition(posicaoInicialX, posicaoInicialY);
        //corteBase.SetPosition(posicaoInicialX, posicaoInicialY);

        imgBase.CenterX = posicaoInicialX;
        imgBase.CenterY = posicaoInicialY;

        corteBase.CenterX = posicaoInicialX;
        corteBase.CenterY = posicaoInicialY;
        // -----------------------------

        // -----------------------------
        // DUPLICAÇÃO HORIZONTAL
        // -----------------------------
        var duplicados = [];

        var atual = doc.CreateShapeRangeFromArray(imgBase, corteBase);

        for (var i = 0; i < quantidadeDeCopiasHorizontal; i++) {

            atual = doc.CreateShapeRangeFromArray(atual.Item(2), atual.Item(1)).Duplicate(deslocamentoHorizontal);

            duplicados.push(atual);
        }
        
        doc.ClearSelection(); //limpa seleção

        // -----------------------------
        // CRIA SHAPERANGE COM TODOS OS PARES
        // -----------------------------
        var rangeSel = host.CreateShapeRange();

        for (var j = 0; j < duplicados.length; j++) {
            rangeSel.Add(duplicados[j].Item(1));
            rangeSel.Add(duplicados[j].Item(2));
            rangeSel.Add(page.Layers.Item("Cut Layer").Shapes.Item(quantidadeDeCopiasHorizontal + 1));
            rangeSel.Add(page.Layers.Item("Camada 1").Shapes.Item(quantidadeDeCopiasHorizontal + 1));
        }

        // -----------------------------
        // DUPLICAÇÃO VERTICAL
        // -----------------------------
        
        for (var i = 1; i <= quantidadeDeCopiasVertical; i++) {
        rangeSel.Duplicate(0, deslocamentoVertical * i);
        }
    }

    if (etiquetaEscolar) {
            // -----------------------------
        // Etiqueta tamanho 10*40mm
        // -----------------------------

        // -----------------------------
        // CONFIGURAÇÕES
        // -----------------------------
        var qtdHorizontal = 25;
        var deslocX = 0.47244094488189;

        var deslocY = [
            1.65354330708661,
            3.30708661417322,
            4.96062992125983
        ];

        // -----------------------------
        // OBJETOS BASE
        // -----------------------------
        var imgBase   = page.Layers.Item("Cut Layer").Shapes.Item(1);
        var corteBase = page.Layers.Item("Camada 1").Shapes.Item(1);

        // -----------------------------
        // DUPLICAÇÃO HORIZONTAL
        // -----------------------------
        var duplicados = [];

        var atual = doc.CreateShapeRangeFromArray(
            imgBase,
            corteBase
        );

        for (var i = 0; i < qtdHorizontal; i++) {

            atual = doc.CreateShapeRangeFromArray(
                atual.Item(2),
                atual.Item(1)
            ).Duplicate(deslocX);

            duplicados.push(atual);
        }

        // -----------------------------
        // LIMPA SELEÇÃO
        // -----------------------------
        doc.ClearSelection();

        // -----------------------------
        // CRIA SHAPERANGE COM TODOS OS PARES
        // -----------------------------
        var rangeSel = host.CreateShapeRange();

        for (var j = 0; j < duplicados.length; j++) {
            rangeSel.Add(duplicados[j].Item(1));
            rangeSel.Add(duplicados[j].Item(2));
            rangeSel.Add(page.Layers.Item("Cut Layer").Shapes.Item(26));
            rangeSel.Add(page.Layers.Item("Camada 1").Shapes.Item(26));
        }

        // -----------------------------
        // DUPLICAÇÃO VERTICAL
        // -----------------------------
        for (var k = 0; k < deslocY.length; k++) {

        var novoRange = host.CreateShapeRange();

        for (var i = 1; i <= rangeSel.Count; i++) {
            var dup = rangeSel.Item(i).Duplicate(0, deslocY[k]);
            novoRange.Add(dup);
        }

        // alert("Script executado com sucesso!");
        }




        // -----------------------------
        // Etiqueta tamanho 20*50mm
        // -----------------------------

        // -----------------------------
        // CONFIGURAÇÕES
        // -----------------------------
        var qtdHorizontal = 13;
        var deslocX = 0.86614285714286;

        var deslocY = [
            2.04724409448819,
            4.09449188818897,
        ];

        // -----------------------------
        // OBJETOS BASE
        // -----------------------------
        var imgBase   = page.Layers.Item("Cut Layer").Shapes.Item(105);
        var corteBase = page.Layers.Item("Camada 1").Shapes.Item(105);

        // -----------------------------
        // DUPLICAÇÃO HORIZONTAL
        // -----------------------------
        var duplicados = [];

        var atual = doc.CreateShapeRangeFromArray(
            imgBase,
            corteBase
        );

        for (var i = 0; i < qtdHorizontal; i++) {

            atual = doc.CreateShapeRangeFromArray(
                atual.Item(2),
                atual.Item(1)
            ).Duplicate(deslocX);

            duplicados.push(atual);
        }

        // -----------------------------
        // LIMPA SELEÇÃO
        // -----------------------------
        doc.ClearSelection();

        // -----------------------------
        // CRIA SHAPERANGE COM TODOS OS PARES
        // -----------------------------
        var rangeSel = host.CreateShapeRange();

        for (var j = 0; j < duplicados.length; j++) {
            rangeSel.Add(duplicados[j].Item(1));
            rangeSel.Add(duplicados[j].Item(2));
            rangeSel.Add(page.Layers.Item("Cut Layer").Shapes.Item(118));
            rangeSel.Add(page.Layers.Item("Camada 1").Shapes.Item(118));
        }

        // -----------------------------
        // DUPLICAÇÃO VERTICAL
        // -----------------------------
        for (var k = 0; k < deslocY.length; k++) {

        var novoRange = host.CreateShapeRange();

        for (var i = 1; i <= rangeSel.Count; i++) {
            var dup = rangeSel.Item(i).Duplicate(0, deslocY[k]);
            novoRange.Add(dup);
        }

        // alert("Script executado com sucesso!");
        }




        // -----------------------------
        // Etiqueta tamanho 30*60mm
        // -----------------------------

        // -----------------------------
        // CONFIGURAÇÕES
        // -----------------------------
        var qtdHorizontal = 8;
        var deslocX = 1.25984285714286;

        var deslocY = [
            2.44094308708661,
        ];

        // -----------------------------
        // OBJETOS BASE
        // -----------------------------
        var imgBase   = page.Layers.Item("Cut Layer").Shapes.Item(147);
        var corteBase = page.Layers.Item("Camada 1").Shapes.Item(147);

        // -----------------------------
        // DUPLICAÇÃO HORIZONTAL
        // -----------------------------
        var duplicados = [];

        var atual = doc.CreateShapeRangeFromArray(
            imgBase,
            corteBase
        );

        for (var i = 0; i < qtdHorizontal; i++) {

            atual = doc.CreateShapeRangeFromArray(
                atual.Item(2),
                atual.Item(1)
            ).Duplicate(deslocX);

            duplicados.push(atual);
        }

        // -----------------------------
        // LIMPA SELEÇÃO
        // -----------------------------
        doc.ClearSelection();

        // -----------------------------
        // CRIA SHAPERANGE COM TODOS OS PARES
        // -----------------------------
        var rangeSel = host.CreateShapeRange();

        for (var j = 0; j < duplicados.length; j++) {
            rangeSel.Add(duplicados[j].Item(1));
            rangeSel.Add(duplicados[j].Item(2));
            rangeSel.Add(page.Layers.Item("Cut Layer").Shapes.Item(155));
            rangeSel.Add(page.Layers.Item("Camada 1").Shapes.Item(155));
        }

        // -----------------------------
        // DUPLICAÇÃO VERTICAL
        // -----------------------------
        for (var k = 0; k < deslocY.length; k++) {

        var novoRange = host.CreateShapeRange();

        for (var i = 1; i <= rangeSel.Count; i++) {
            var dup = rangeSel.Item(i).Duplicate(0, deslocY[k]);
            novoRange.Add(dup);
        }

        // alert("Script executado com sucesso!");
        }
    }


    
    
    if (corteRedondo) {
        
        
        
    alert("O script rodou sem erros chefia");
    }
}

main();