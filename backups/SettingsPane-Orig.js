// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

/// <reference path="../App.js" />

(function () {
    "use strict";
  
    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
      $(document).ready(function () {
          app.initialize();

          _settings = Office.context.roamingSettings;
          
          var customer = _settings.get("customer")
          $('#customerToAdd').val(customer);

          $('#addCustomer').click(addCustomer);
      });
    };
    
    function addCustomer() {
      var customerToAdd = $('#customerToAdd').val();

      _settings.set("customer", customerToAdd);

      return new Promise((resolve, reject) => {  
          // Fake success  
          setTimeout(() => {  
            resolve("success");  
          }, 1000);
      });    
    }
  }
);