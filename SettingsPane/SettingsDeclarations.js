
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
            setBillingRates(127);
        });
    };
})();

var runningTotalHours;
var runningTotalAmount;


function setBillingRates(rate) {

  runningTotalHours = 0;
  runningTotalAmount = 0;

  var xhr = [
    setMeetingsRate(rate),
    setMeetingsRate(rate),
    setMeetingsRate(rate)
  ];

  $.when(xhr[0], xhr[1], xhr[2]).then(function () {
    //showGrandTotal();
  });
}


// Constructs the meetings table and calculated the total
// billing amount for each item and for all meetings.
function setMeetingsRate(rate) {
    return $.getJSON("../assets/customerList.json", function (data) {
        var jsonData = data.Customers;

        var dataTable = $("#app-Meetings-table");
        dataTable.html("");

        var headerRow = $('<div />');
        headerRow.append(makeHeaderCell("Name", "5"));
        headerRow.append(makeHeaderCell("RMID", "5"));
        headerRow.append(makeHeaderCell("SOW", "1", "true"));

        dataTable.append(headerRow);

        var totalHours = 0;
        var totalAmount = 0;

        for (var i in jsonData) {
            var dataRow = $("<div />", {
                "class": "ms-Grid-row app-Grid-row"
            });
            dataRow.append(makeRowCell(jsonData[i].Name, "5"));
            dataRow.append(makeRowCell(jsonData[i].RMID, "5"));
            dataRow.append(makeRowCell(jsonData[i].SOW_Path, "1", "true"));

            totalHours += Number(jsonData[i].Hours);
            totalAmount += rate * (jsonData[i].Hours);

            dataTable.append(dataRow);
        }
        

        runningTotalHours += totalHours;
        runningTotalAmount += totalAmount;
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



// // Creates the row that displays the grand total for the page.
// function showGrandTotal() {
//     var totalTable = $("#app-Running-total");
//     totalTable.html("");

//      var totalRow = $("<div />", {
//          "class": "app-Title-bar ms-bgColor-themeDarker ms-fontColor-themeLighter ms-font-xxl ms-fontWeight-semibold"
//      });

//      totalRow.append($("<div />", {
//          "class": "app-Cell-right ms-Grid-col ms-u-lg10",
//          "html": "Grand total:"
//      }));

//      totalRow.append($("<div />", {
//          "class": "app-Cell-right ms-Grid-col ms-u-lg1",
//          "html": runningTotalHours
//      }));

//      totalRow.append($("<div />", {
//          "class": "app-Cell-right ms-Grid-col ms-u-lg1",
//          "html": runningTotalAmount
//      }));

//     totalTable.append(totalRow);
// };


// function completeEvent(event) {
//     if (event) {
//         event.completed(true);
//     }
// };

