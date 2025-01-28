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
  
            $('#save').click(saveAgenda);
        });
    };
    
    function saveAgenda(event){

        var text = $('#subjectToSave').val();
        _settings.set("subject", text);

        var html = $('#bodyToSave').val();
        _settings.set("body", html);        
    }
    
})();