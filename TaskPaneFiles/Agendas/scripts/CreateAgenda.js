/* (function () {
    "use strict"
    // The initialze function is run each time the page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            //adds listener to Phases for any change
            setPhasesListener();

            //adds listener to submit button for any change
            setGenerateAgendaListener();

            //populates our default list values
            getCustomersToList();
            getPhasesToList();
            getAgendasToList();
        });
    };
})(); */


function setPhasesListener() {
    return $("#app-Phase-dropdown").on("change", getAgendasToList);
}


function setGenerateAgendaListener() {
    return $("#app-generateAgenda-button").on("click", generateAgenda);
}


function getCustomersToList() {
    return $.getJSON("../../Assets/Json/customerList.json", function (data) {
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
    return $.getJSON("../../Assets/Json/phaseList.json", function (data) {
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
    return $.getJSON("../../Assets/Json/agendaList.json", function (data) {
        var jsonData = data.Agendas;

        var selectedPhase = $("#app-Phase-dropdown").val();

        var dataTable = $("#app-Agenda-dropdown");
        dataTable.html("");


        for (var i in jsonData) {

            if (jsonData[i].name == selectedPhase) {


                for (var j in jsonData[i].type) {

                    var dataOption = $("<option />", {
                        "class": "ms-Dropdown-item"
                    });

                    dataOption.append(jsonData[i].type[j]);


                    dataTable.append(dataOption);
                }

            }

        }

    });
}


function generateAgenda() {
    return $.getJSON("../../Assets/Json/agendaDetails.json", function (data) {
        var jsonData = data.Agendas;

        var selectedCustomer = $("#app-Customer-dropdown").val();
        var selectedPhase = $("#app-Phase-dropdown").val();
        var selectedAgenda = $("#app-Agenda-dropdown").val();

        var agendaName = `${selectedPhase} - ${selectedAgenda}`;
        
        for (var i in jsonData) {
            
            //check for matching agenda in agenda details
            if (agendaName == jsonData[i].name) {

                //set the text subject
                var agendaSubject = jsonData[i].subject

                //Office.context.mailbox.item.subject.setAsync(`${selectedCustomer} - ${agendaSubject}`, function (asyncResult) { });
                setTextToSubject(`${selectedCustomer} - ${agendaSubject}`);

                //replace agenda with SOW
                var bodyAgenda = jsonData[i].bodyAgenda
                
                setSOWToBody(bodyAgenda)

                //set the HTML body
                var agendaBody = jsonData[i].bodyIntro + jsonData[i].bodyObjective + jsonData[i].bodyAgenda + jsonData[i].bodyClosing

                setHTMLToBody(agendaBody);


            }
        }
    });
}


function setTextToSubject(text) {

    return new Promise(function (resolve, reject) {
        try {

            //_mailbox.item.subject.setAsync(
            Office.context.mailbox.item.subject.setAsync(
                text,
                function (asyncResult) {
                    //statusUpdate(asyncResult,"Insert");
                    resolve();
                }
            );
        }
        catch (error) {
            reject();
        }
    })
}


function setSOWToBody(html) {

    
}


function setHTMLToBody(html) {

    return new Promise(function (resolve, reject) {
        try {

            //_mailbox.item.body.setSelectedDataAsync(
            Office.context.mailbox.item.body.setSelectedDataAsync(
                html,
                { coercionType: Office.CoercionType.Html },
                function (asyncResult) {
                    //statusUpdate(asyncResult,"Insert");
                    resolve();
                }
            );
        }
        catch (error) {
            reject();
        }
    })
}