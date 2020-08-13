$('#changeColorButton').click(function () {
    var color = $('#changeColorSelect').val();
    console.log(color)
    $('body').css("background-color", color);
    localStorage.setItem('background', color);
});