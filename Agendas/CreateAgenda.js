
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
            getAgendasToList();
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
        var jsonData = data.Processes;

        var dataTable = $("#app-Process-dropdown");
        dataTable.html("");

        for (var i in jsonData) {

            var dataOptgroup = $("<optgroup />", {
                "label": jsonData[i].Name,
                "class": "ms-Dropdown-label"                
            });

            //dataOptgroup.append(jsonData[i].Name);
            
            for (var j in jsonData[i].type){

                var dataOption = $("<option />",{
                    "class": "ms-Dropdown-item"
                });

                dataOption.append(jsonData[i].type[j]);

                dataOptgroup.append(dataOption);
            }


            dataTable.append(dataOptgroup);
        }

    });
};
function getAgendasToList() {
    return $.getJSON("../Assets/agendaList.json", function (data) {
        var jsonData = data.Agendas;

        var dataTable = $("#app-Agenda-dropdown");
        dataTable.html("");

        for (var i in jsonData) {

            var dataOptgroup = $("<optgroup />", {
                "label": jsonData[i].Name,
                "class": "ms-Dropdown-label"                
            });

            //dataOptgroup.append(jsonData[i].Name);
            
            for (var j in jsonData[i].type){

                var dataOption = $("<option />",{
                    "class": "ms-Dropdown-item"
                });

                dataOption.append(jsonData[i].type[j]);

                dataOptgroup.append(dataOption);
            }


            dataTable.append(dataOptgroup);
        }

    });
};



// function completeEvent(event) {
//     if (event) {
//         event.completed(true);
//     }
// };

