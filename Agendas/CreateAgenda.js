
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
            getProcessesToList();
        });
    };
})();


function getCustomersToList() {
    return $.getJSON("../Assets/customerList.json", function (data) {
        var jsonData = data.Customers;

        var dataTable = $("#app-Customer-dropdown");
        dataTable.html("");

        for (var i in jsonData) {
            var dataRow = $("<option />", {
                "class": "ms-Dropdown-item"
            });
            dataRow.append(jsonData[i].Name);

            dataTable.append(dataRow);
        }

    });
};

function getProcessesToList() {
    return $.getJSON("../Assets/processList.json", function (data) {
        var jsonData = data.processes

        var dataTable = $("#app-Process-dropdown");
        dataTable.html("");

        for (i in jsonData.process) {
            var dataGroup = $("<optgroup />", {
                "class": "ms-Dropdown-items"
            });

            dataGroup.append(jsonData.process[i]);


            for (j in jsonData.type){
                var dataOption = $("<option />", {
                    "class": "ms-Dropdown-item"
                });

                dataGroup.append(jsonData.type[i]);
            }
        }


    });
};



// function completeEvent(event) {
//     if (event) {
//         event.completed(true);
//     }
// };

