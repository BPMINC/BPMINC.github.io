// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

/// <reference path="../App.js" />

(function () {
  'use strict';

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      app.initialize();

      $('#insertDefault').click(insertDefault);
      $('#insertMsg1').click(insertMsg1);
      $('#insertCustom').click(insertCustom);
    });
  };
  
  function insertText(textToInsert) {
    // Insert as plain text (CoercionType.Text)
    Office.context.mailbox.item.body.setSelectedDataAsync(
      textToInsert, 
      { coercionType: Office.CoercionType.Text }, 
      function (asyncResult) {
        // Display the result to the user
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
          app.showNotification("Success", "\"" + textToInsert + "\" inserted successfully.");
        }
        else {
          app.showNotification("Error", "Failed to insert \"" + textToInsert + "\": " + asyncResult.error.message);
        }
      });
  }

  function insertDefault() {
    insertText("Inserted by the Add-in Command Demo add-in.");
  }
  
  function insertMsg1() {
    insertText("Hello World!");
  }
  
  function insertCustom() {
    var textToInsert = $('#textToInsert').val();
    insertText(textToInsert);
  }
})();