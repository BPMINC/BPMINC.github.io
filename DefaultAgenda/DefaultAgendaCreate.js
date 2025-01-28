// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

/// <reference path="../App.js" />

let _mailbox
let _settings;

(function () {
  "use strict";

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      app.initialize();
      
      _mailbox = Office.context.mailbox;
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

        console.log("saved4")
        Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
          type: "informationalMessage",
          icon: "icon-16",
          message: "Save successful",
          persistent: false
        });        
      }
      else {
        Office.context.mailbox.item.notificationMessages.addAsync("error", {
          type: "errorMessage",
          message: "Save Failed - " + asyncResult.error.message,
          persistent: false
        }); 
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