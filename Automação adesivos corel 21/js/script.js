const formulario = document.getElementById("formulario");

formulario.innerHTML = `

<div class="section">

<h3>Tipo de Corte</h3>

<div class="item">

<label>

<input type="radio" name="corte" checked>

Quadrado

</label>

</div>

<div class="item">

<label>

<input type="radio" name="corte">

Redondo

</label>

</div>

<div class="item">

<label>

<input type="radio" name="corte">

Personalizado

</label>

</div>

<div class="item">

<label>

<input type="radio" name="corte">

Etiqueta Escolar

</label>

</div>

</div>

<div class="section">

<h3>Tamanho</h3>

<div class="item">

<label>Largura (mm)</label>

<input type="number" value="25">

</div>

<div class="item">

<label>Altura (mm)</label>

<input type="number" value="25">

</div>

</div>

<div class="section">

<h3>Quantidade de Cópias</h3>

<div class="item">

<label>Horizontal</label>

<input type="number" value="1">

</div>

<div class="item">

<label>Vertical</label>

<input type="number" value="1">

</div>

</div>

<div class="section">

<h3>Posição Inicial</h3>

<div class="item">

<label>X (mm)</label>

<input type="number" value="0">

</div>

<div class="item">

<label>Y (mm)</label>

<input type="number" value="0">

</div>

</div>

`;

document
    .getElementById("btnGerar")
    .addEventListener("click", ()=>{

        alert("Em breve vamos gerar o script.");

});