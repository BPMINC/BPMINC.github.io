




/* /// <reference path="../App.js" /> */


let _mailbox;
let _settings;

/*
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
} */


/* (function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        //$(document).ready(function () {
            //app.initialize();
       // });
    }; */

/*     function insertDefaultAgenda(event) {

        var subject = _settings.get("subject");
        console.log(subject + " - sub7")

        //setTextToSubject(subject);

        //var body = _settings.get("body");
        //console.log(body + " - body7");

        //await setHTMLToBody(body, event);

        event.completed();
    }
    
    
    
    function setTextToSubject(text) {

        _mailbox.item.subject.setAsync(text, 

            function (asyncResult){
                // Display the result to the user
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    app.showNotification("Success", "\"" + text + "\" inserted successfully.");       
                }
                else {
                    app.showNotification("Error", "Failed to insert \"" + textToInsert + "\": " + asyncResult.error.message);
                }    
            }
        ); 
    } 
    
    function setHTMLToBody(html, event) {
        _mailbox.item.body.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, 

            function (asyncResult){
                // Display the result to the user
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _mailbox.item.notificationMessages.replaceAsync("status", {
                    type: "informationalMessage",
                    icon: "icon-16",
                    message: "Save successful",
                    persistent: false
                    });        
                }
                else {
                    _mailbox.item.notificationMessages.replaceAsync("error", {
                    type: "errorMessage",
                    message: "Save Failed - " + asyncResult.error.message,
                    persistent: false
                    }); 
                }   
            }
        );
    }  */

/* })(); */