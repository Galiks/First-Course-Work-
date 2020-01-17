window.onload = Init

function Init() {
    document.getElementById("iframe_redactor").contentWindow.document.designMode = "On";
}

function doStyle(style) {
    document.getElementById("iframe_redactor").contentWindow.document.execCommand(style, false, null);
}

function doURL() {
    var image = prompt("Enter a URL:", "http://");
    if ((image != null) && (image != "")) {
        document.getElementById("iframe_redactor").contentWindow.document.execCommand("insertImage", false, image);
    }
}

class createTable {
    constructor(rows, cols) {
        this.rowNum = rows;
        this.colNum = cols;
        this.table = "<table>";
    }

    make() {
        for (let i = 0; i < this.rowNum; i++) {
            this.table += "<tr>";
            for (let j = 0; j < this.colNum; j++) {
                this.table += `<td>${i}.${j}</td>`;
            }
            this.table += "</tr>"
        }
        this.table += "</table>"
        return this;
    }
}

function addTable() {
    var ct = new createTable(row.value, col.value).make().table;
    document.getElementById("iframe_redactor").contentWindow.document.execCommand("insertHTML", true, ct);
}