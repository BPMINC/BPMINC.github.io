
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
            getCustomersToList();
        });
    };
})();


// Constructs the meetings table and calculated the total
// billing amount for each item and for all meetings.
function getCustomersToList() {
    return $.getJSON("../Assets/customerList.json", function (data) {
        var jsonData = data.Customers;

        var dataTable = $("#app-Customers-dropdown");
        dataTable.html("");

        for (var i in jsonData) {
            var dataRow = $("<div />", {
                "class": "ms-Dropdown-item"
            });
            dataRow.append(jsonData[i].Name);

            dataTable.append(dataRow);
        }

    });
};


function makeDropdownItem(text) {
    var cssClass = "ms-Dropdown-item ms-u-md4 ms-u-lg3" ;


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

