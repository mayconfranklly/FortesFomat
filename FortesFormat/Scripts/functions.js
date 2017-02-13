$('#btnSubmit').on('click', function (e) {
    if ($('.thumbnail').val() == "") {
        alert('Selecione um arquivo.');
        e.preventDefault();
    }
});