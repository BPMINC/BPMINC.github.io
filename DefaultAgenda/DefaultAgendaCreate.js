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

            var body = _settings.get("subject")
            $('#bodyToSave').val(body);

            $('#save').click(saveAgenda);
        });
    };
    
    async function saveAgenda(event){
        await saveSubject();
        await saveBody();  
        
        await _settings.saveAsync();

        event.completed()
    }

    async function saveSubject(){
        var text = $('#subjectToSave').val();
        _settings.set("subject", text);

        return new Promise((resolve, reject) => {  
            // Fake success  
            setTimeout(() => {  
              resolve("success");  
            }, 1000);
        });
    }

   async function saveBody(){
        var html = $('#bodyToSave').val();
        _settings.set("body", html);  

        return new Promise((resolve, reject) => {  
            // Fake success  
            setTimeout(() => {  
              resolve("success");  
            }, 1000);
        });
    }
    
})();