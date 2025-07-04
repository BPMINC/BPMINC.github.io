// (function () {
//     "use strict"
//     // The initialze function is run each time the page is loaded.
//     Office.initialize = function (reason) {
//         $(document).ready(function () {
//             getCustomersToTable();
//         });
//     };
// })();


//Purpose: Defines an expression (not a definition) to be evaluated
"use strict"

// The initialize function/expression is run each time the page is loaded.
Office.initialize = function () {

    $(document).ready(function () {
        getCustomersToTable();
    });

};