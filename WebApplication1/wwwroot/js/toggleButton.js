var elem = document.getElementById("switch");

if (elem !== null) {
    elem.onclick = function () {
        document.getElementById("slider").classList.toggle("change");
        document.getElementById("circle").classList.toggle("change");
        document.getElementById("circle-content").classList.toggle("change");
    };
}