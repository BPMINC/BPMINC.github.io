// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

let _mailbox;
let _settings;
let _customerName;

Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
}

function addDefaultMsg(event) {
  setHTMLToSubject("Default Agenda Subject", "icon-16", event);
  setHTMLToBody("This is the default agenda text<br/>", "icon-16", event);
}

function addP2PMsg(event) {
  _settings.set("customerName", "OneWebTech");
  _settings.saveAsync();

  _customerName = _settings.get("customerName");

  setHTMLToSubject(_customerName + " - P2P Requirements Gathering", "icon-16", event);
  setHTMLToBody("<b><i>Meeting Objective</i></b><br/><br/>\
    The objective of this session is for our team to gather \
    a solid understanding of your AP processes from vendor \
    creation, vendor bills and associated approvals, vendor \
    payments and advanced electronic payments and expense reporting\
    <br/><br/><br/><b><i>Meeting Agenda</i></b><br/>\
    <ul><li>Vendor Master</li><li>Employee Master</li>\
    <li>Vendor Bills</li><li>Vendor Payments</li>\
    <li>Expense Reports</li><li>Fixed Assets</li></ul>\
    ", "icon-16", event);
}


async function setHTMLToSubject(text, icon, event) {
  _mailbox.item.subject.setAsync(text, 
    function (asyncResult){
      if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        _mailbox.item.notificationMessages.replaceAsync("status", {
          type: "informationalMessage",
          icon: icon,
          message: "Success",
          persistent: false
        });
      }
      else {
        _mailbox.item.notificationMessages.addAsync("error", {
          type: "errorMessage",
          message: "Failed - " + asyncResult.error.message,
          persistent: false
        });
      }
      
      return new Promise(resolve, reject);

    }
  );    
  event.completed();
} 

async function setHTMLToBody(text, icon, event) {
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
    }
  );
  event.completed();
} 