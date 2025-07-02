
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
            getPhasesToList();
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


function getPhasesToList() {
    return $.getJSON("../Assets/phaseList.json", function (data) {
        var jsonData = data.Phases;

        var dataTable = $("#app-Phase-dropdown");
        dataTable.html("");

        for (var i in jsonData) {

            var dataOption = $("<option />", {
                "class": "ms-Dropdown-item"                
            });

            dataOption.append(jsonData[i]);


            dataTable.append(dataOption);
        }

    });
};


function getAgendasToList() {
    return $.getJSON("../Assets/agendaList.json", function (data) {
        var jsonData = data.Agendas;

        var dataTable = $("#app-Agenda-dropdown");
        dataTable.html("");

        var selectedVal = $("#app-Phase-dropdown").val();

        for (var i in jsonData) {

            if (jsonData[i].name == selectedVal) {


                for (var j in jsonData[i].type)

                var dataOption = $("<option />", {
                    "class": "ms-Dropdown-item"                
                });
    
                dataOption.append(jsonData[i].type[j]);
    
    
                dataTable.append(dataOption);
            }

        }

    });
}



// function completeEvent(event) {
//     if (event) {
//         event.completed(true);
//     }
// };

