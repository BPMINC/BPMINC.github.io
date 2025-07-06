//Purpose: Defines an expression (not a definition) to be evaluated
"use strict"

// The initialize function/expression is run each time the page is loaded.
Office.initialize = function () {

    $(document).ready(function () {

        //adds an onChange event to Phases html
        setPhasesOnChange();

        //adds an onClick event to submit button html
        setCreateAgendaOnClick();

        //populates our default list values
        getCustomersToList();
        getPhasesToList();
        getAgendasToList();
    });

};