let _mailbox;
let _settings;
let _subject;
let _body;

Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
}

function insertDefaultAgenda(event) {

  _subject = _settings.get("subject");
  setTextToSubject("somesubject", "icon-16", event);

  _body = _settings.get("body");
  setHTMLToBody("someBody", "icon-16", event);
}



async function setTextToSubject(text, icon, event) {

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

function setHTMLToBody(html, icon, event) {
  Office.context.mailbox.item.body.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, 
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