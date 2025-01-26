// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

Office.initialize = function () {
}

function addDefaultMsg(event) {
  setHTMLToBody("<b>This is the test agenda text</b><br/>", "blue-icon-16", event);
}

function addP2PMsg(event) {
  setHTMLToBody("<b><i>Meeting Objective</i></b><br/><br/>\
    The objective of this session is for our team to gather a solid understanding of your AP processes from vendor creation\
    , vendor bills and associated approvals, vendor payments and advanced electronic payments and expense reporting\
    <br/><br/><br/><b><i>Meeting Agenda</i></b><br/>\
    <ul><li>Vendor Master</li><li>Employee Master</li><li>Vendor Bills</li></ul>", "blue-icon-16", event);
}


function setHTMLToBody(text, icon, event) {
  Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Html }, 
    function (asyncResult){
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
      event.completed();
    });
} 