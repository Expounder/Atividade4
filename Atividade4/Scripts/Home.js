$("#Bt_Enviar").click(function () {

    var arquivo = new FormData();
    arquivo.append($("#resultado")[0].files[0].name, $("#resultado")[0].files[0])
    SendFile(arquivo);
});
    

function SendFile(arquivo) {
    $.ajax({
        url: '/Home/Envio',
        data: arquivo,
        processData: false,
        contentType: false,
        type: 'POST',
        success: function (data) {
            if (data.success) {
                alert("Os emails foram enviados a seus destinatários")
            }
            else {
                alert("Falha no Upload")
            }
        }
    });
}
