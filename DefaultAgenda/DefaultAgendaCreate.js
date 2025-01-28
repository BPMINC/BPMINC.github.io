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
        
        await saveSubject();
        await saveBody();  
        
        await _settings.saveAsync(updateNotification);

        event.completed()
    }

    async function saveSubject(){
        var text = $('#subjectToSave').val();
        _settings.set("subject", text);

        return new Promise((resolve, reject) => {  
            // Fake success  
              resolve("success");
        });
    }

   async function saveBody(){
        var html = $('#bodyToSave').val();
        _settings.set("body", html);  

        return new Promise((resolve, reject) => {  
            // Fake success   
              resolve("success");  
        });        
    }

    function updateNotification(asyncResult){
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
          Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
            type: "informationalMessage",
            icon: icon,
            message: "Success",
            persistent: false
          });
        }
        else {
          Office.context.mailbox.item.notificationMessages.addAsync("error", {
            type: "errorMessage",
            message: "Failed - " + asyncResult.error.message,
            persistent: false
          });
        }
        
        return new Promise(resolve, reject);
  
    }
    
})();