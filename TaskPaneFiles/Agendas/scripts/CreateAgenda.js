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

            if (jsonData[i].Name == selectedPhase) {

                for (var j in jsonData[i].Type) {

                    var dataOption = $("<option />", {
                        "class": "ms-Dropdown-item"
                    });

                    dataOption.append(jsonData[i].Type[j]);

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
    createAttendees(selectedCustomer);
    createSubject(selectedCustomer, selectedPhase, selectedAgenda);
    createBody(selectedCustomer, selectedPhase, selectedAgenda);

}


function createCategory(customer) {
    return $.getJSON("../../Assets/Json/customerList.json", function (data) {
        var jsonData = data.Customers;

        for (var i in jsonData) {

            if (customer == jsonData[i].Name) {

                var projectId = jsonData[i].RMID
                var category = `${customer} - ${projectId}`

                addTextToCategories(category);
            }
        }
    });
}


function addTextToCategories(text) {
    return new Promise(function (resolve, reject) {
        try {

            Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

                    var categoryList = asyncResult.value;

                    if (categoryList && categoryList.length > 0) {

                        for (var i in categoryList) {

                            var category = [categoryList[i].displayName];

                            Office.context.mailbox.item.categories.removeAsync(category, function (asyncResult) {
                                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

                                    console.log("categories - removeAsync success")

                                } else {

                                    console.log("categories - removeAsync failed");
                                }
                            });
                        }
                    } else {
                        console.log("categories - no categories to remove");
                    }

                } else {
                    console.log("categories - getAsync failed")
                }
            });

            Office.context.mailbox.item.categories.addAsync([text], function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

                    //statusUpdate(asyncResult,"Insert");

                    console.log("categories - addASync success");

                    resolve();

                }
                else {
                    console.log("categories - addASync failed");
                }
            });
        }
        catch (error) {

            console.log("categories - promise failed:" + error);
            reject();
        }
    });
}


function createAttendees(customer) {
    return $.getJSON("../../Assets/Json/customerList.json", function (data) {
        var jsonData = data.Customers;

        for (var i in jsonData) {

            if (customer == jsonData[i].Name) {

                var attendeeList = [{emailAddress: "JosephSmith@bpmcpa.com", displayName: "Joseph Smith"}]

                setTextToAttendees(attendeeList);

            }
        }
    });
}


function setTextToAttendees(text) {
    return new Promise(function (resolve, reject) {
        try {

/*             Office.context.mailbox.item.requiredAttendees.getAsync(function (asyncResult){
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

                    var attendees = asyncResult.value;

                    console.log(attendees);
                    console.log(text);
                } else{
                    console.log("attendees - getAsync failed")
                }
            }); */


            Office.context.mailbox.item.requiredAttendees.setAsync(text, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

                    //statusUpdate(asyncResult,"Insert");

                    console.log("attendees - setASync success");

                    resolve();

                }
                else {
                    console.log("attendees - setASync failed");
                }
            });
        }
        catch (error) {

            console.log("attendees - promise failed: " + error);
            reject();
        }
    });
}


function createSubject(customer, phase, agenda) {
    return $.getJSON("../../Assets/Json/agendaDetails.json", function (data) {

        var jsonData = data.Agendas;
        var agendaName = `${phase} - ${agenda}`;

        for (var i in jsonData) {

            //check for matching agenda in agenda details json
            if (agendaName == jsonData[i].Name) {

                var subject = jsonData[i].Subject
                setTextToSubject(`${customer} - ${subject}`);
            }
        }
    });
}


function setTextToSubject(text) {
    return new Promise(function (resolve, reject) {
        try {
            Office.context.mailbox.item.subject.setAsync(text, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

                    //statusUpdate(asyncResult,"Insert");

                    console.log("subject - setASync success");

                    resolve();

                }
                else {
                    console.log("subject - setASync failed");
                }
            });
        }
        catch (error) {

            console.log("subject - promise failed: " + error);

            reject();
        }
    });
}



function createBody(customer, phase, agenda) {
    return $.getJSON("../../Assets/Json/agendaDetails.json", function (data) {
        var jsonData = data.Agendas;


        var agendaName = `${phase} - ${agenda}`;

        for (var i in jsonData) {

            //check for matching agenda in agenda details json
            if (agendaName == jsonData[i].Name) {


                //set the body HTML
                var body = jsonData[i].BodyIntro + jsonData[i].BodyObjective + jsonData[i].BodyAgenda + jsonData[i].BodyClosing

                setHTMLToBody(body);
            }
        }
    });
}


function setHTMLToBody(html) {
    return new Promise(function (resolve, reject) {
        try {
            Office.context.mailbox.item.body.setAsync(html, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

                    //statusUpdate(asyncResult,"Insert");

                    console.log("body - setASync success");

                    resolve();

                }
                else {
                    console.log("body - setASync failed");
                }
            });
        }
        catch (error) {

            console.log("subject - promise failed: " + error);

            reject();
        }
    });

}