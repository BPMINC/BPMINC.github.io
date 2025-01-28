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
  
  async function saveAgenda(event){
      
    saveSubject();
    saveBody();  
    
    await _settings.saveAsync(statusUpdate);

    event.completed();
  }

  function saveSubject(){
      var text = $('#subjectToSave').val();
      _settings.set("subject", text);;
  }

  function saveBody(){
      var html = $('#bodyToSave').val();
      _settings.set("body", html);       
  }

  function statusUpdate(asyncResult){
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
        type: "informationalMessage",
        message: "Success",
        persistent: false
      });
    }
    else {
      Office.context.mailbox.item.notificationMessages.replaceAsync("error", {
        type: "errorMessage",
        message: "Failed - " + asyncResult.error.message,
        persistent: false
      });
    }
  }
    
})();