
//grabs the customerList json and formats it into a css/ html grid for display
function getCustomersToTable() {
    return $.getJSON("../../Assets/Json/customerList.json", function (data) {
        var jsonData = data.Customers;

        var dataTable = $("#app-Customers-table");
        dataTable.html("");

        for (var i in jsonData) {
            var dataRow = $("<div />", {
                "class": "ms-Grid-row app-Grid-row"
            });
            dataRow.append(makeRowCell(jsonData[i].Name, "3"));
            dataRow.append(makeRowCell(jsonData[i].RMID, "2"));
            dataRow.append(makeRowCell(jsonData[i].SOW_Path, "5", "true"));

            dataTable.append(dataRow);
        }

    });
};


function makeRowCell(text, width, right) {
    var cssClass = "ms-Grid-col ms-u-md4 ms-u-lg" + width;

    if (right) {
        cssClass += " app-Cell-right";
    }

    return $("<div />", {
        "class": cssClass,
        "html": text
    });
}