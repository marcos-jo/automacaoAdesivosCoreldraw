try {

    var fso = new ActiveXObject("Scripting.FileSystemObject");

    alert("FileSystemObject OK");

}
catch(e){

    alert("Erro:\n" + e.message);

}