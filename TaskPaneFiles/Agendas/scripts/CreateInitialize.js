//Purpose: Defines an expression (not a definition) to be evaluated
"use strict"

// The initialize function/expression is run each time the page is loaded.
Office.initialize = function () {

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