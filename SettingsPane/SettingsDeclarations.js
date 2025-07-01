
// Office.initialize = function () {

//   "use strict"

//   $(document).ready(function () {

//     setBillingRates(127);
//   });

// }


(function () {
    "use strict"
    // The initialze function is run each time the page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            getCustomers();
        });
    };
})();


// Constructs the meetings table and calculated the total
// billing amount for each item and for all meetings.
function getCustomers() {
    return $.getJSON("../Assets/customerList.json", function (data) {
        var jsonData = data.Customers;

        var dataTable = $("#app-Meetings-table");
        dataTable.html("");

        var headerRow = $('<div />');
        headerRow.append(makeHeaderCell("Name", "3"));
        headerRow.append(makeHeaderCell("RMID", "2"));
        headerRow.append(makeHeaderCell("SOW", "5", "true"));

        dataTable.append(headerRow);



        for (var i in jsonData) {
            var dataRow = $("<div />", {
                "class": "ms-Grid-row app-Grid-row"
            });
            dataRow.append(makeRowCell(jsonData[i].Name, "3"));
            dataRow.append(makeRowCell(jsonData[i].RMID, "2"));
            dataRow.append(makeRowCell(jsonData[i].SOW_Path, "5", "true"));

   

            dataTable.append(dataRow);
        }
        

        console.log(jsonData);

    });
};


function makeHeaderCell(text, width, right) {
    var cssClass = "ms-Grid-col ms-fontColor-themeDark ms-font-l ms-u-lg" + width;

    if (right) {
        cssClass += " app-Cell-right";
    }

    return $("<div />", {
        "class": cssClass,
        "html": text
    })
};

// Creates the HTML for displaying a table cell.
function makeRowCell(text, width, right) {
    var cssClass = "ms-Grid-col ms-u-md4 ms-u-lg" + width;

    if (right) {
        cssClass += " app-Cell-right";
    }

	return $("<div />", {
        "class" : cssClass,
        "html"  : text
    });    
}


// function completeEvent(event) {
//     if (event) {
//         event.completed(true);
//     }
// };

