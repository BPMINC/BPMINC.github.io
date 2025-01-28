// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

/// <reference path="../App.js" />

let _settings;

(function () {
  "use strict";

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      app.initialize();
      
      _settings = Office.context.roamingSettings;

      var subject = _settings.get("subject")
      $('#subjectToSave').val(subject);

      var body = _settings.get("body")
      $('#bodyToSave').val(body);

      $('#save').click(saveAgenda);
    });
  };
  
  function saveAgenda(){

    saveSubject();
    saveBody();  
    
    _settings.saveAsync(function (asyncResult) {
      // Display the result to the user
      if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        console.log(`saved2`);
        Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
          type: "informationalMessage",
          message: "Success",
          persistent: false
        });        
        /* app.showNotification("Success", "saved successfully"); */
      }
      else {
        console.log(`not saved`);
        /* app.showNotification("Error", "Failed to save: " + asyncResult.error.message); */
      }
    })
  }

  function saveSubject(){
      var text = $('#subjectToSave').val();
      _settings.set("subject", text);;
  }

  function saveBody(){
      var html = $('#bodyToSave').val();
      _settings.set("body", html);       
  }
    
})();