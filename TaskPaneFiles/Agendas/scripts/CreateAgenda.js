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


function setPhasesOnChange() {
    return $("#app-Phase-dropdown").on("change", getAgendasToList);
}


function setCreateAgendaOnClick() {
    return $("#app-CreateAgenda-button").on("click", createAgenda);
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


function createAgenda() {

    var selectedCustomer = $("#app-Customer-dropdown").val();
    var selectedPhase = $("#app-Phase-dropdown").val();
    var selectedAgenda = $("#app-Agenda-dropdown").val();

    createCategory(selectedCustomer);
    createAttendees();
    createSubject(selectedCustomer, selectedPhase, selectedAgenda);
    createBody(selectedCustomer, selectedPhase, selectedAgenda);

}


function createCategory(customer) {
    return $.getJSON("../../Assets/Json/customerList.json", function (data) {
        var jsonData = data.Customers;

        for (var i in jsonData) {

            if (customer == jsonData[i].name) {

                var projectId = jsonData[i].RMID;
                var category = `${customer} - ${projectId}`;

                console.log(category);

                addTextToCategories(category);

            }

        }

    });
}


function addTextToCategories(text) {
    return new Promise(function (resolve, reject) {
        try {
            Office.context.mailbox.item.categories.addAsync(
                [text], //must be set inside an array for addAsync to work
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


function createAttendees(customer) {
    return $.getJSON("../../Assets/Json/customerList.json", function (data) {
        var jsonData = data.Customers;

        for (var i in jsonData) {

            if (customer == jsonData[i].name) {

                var category = `${customer} - ${jsonData[i].RMID}`

                addTextToAttendees(category);

            }

        }

    });
}


function addTextToAttendees(text) {
    return new Promise(function (resolve, reject) {
        try {
            Office.context.mailbox.item.categories.addAsync(
                ["josephsmith@bpmcpa.com"], //must be set inside an array for addAsync to work
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


function createSubject(customer, phase, agenda) {
    return $.getJSON("../../Assets/Json/agendaDetails.json", function (data) {
        var jsonData = data.Agendas;


        var agendaName = `${phase} - ${agenda}`;

        for (var i in jsonData) {

            //check for matching agenda in agenda details json
            if (agendaName == jsonData[i].name) {


                var subject = jsonData[i].subject

                //Office.context.mailbox.item.subject.setAsync(`${selectedCustomer} - ${agendaSubject}`, function (asyncResult) { });
                setTextToSubject(`${customer} - ${subject}`);
            }
        }
    });
}


function setTextToSubject(text) {
    return new Promise(function (resolve, reject) {
        try {
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



function createBody(customer, phase, agenda) {
    return $.getJSON("../../Assets/Json/agendaDetails.json", function (data) {
        var jsonData = data.Agendas;


        var agendaName = `${phase} - ${agenda}`;

        for (var i in jsonData) {

            //check for matching agenda in agenda details json
            if (agendaName == jsonData[i].name) {


                //set the body HTML
                var body = jsonData[i].bodyIntro + jsonData[i].bodyObjective + jsonData[i].bodyAgenda + jsonData[i].bodyClosing

                setHTMLToBody(body);
            }
        }
    });
}


function setHTMLToBody(html) {
    return new Promise(function (resolve, reject) {
        try {
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