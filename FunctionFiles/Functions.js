
let _mailbox
let _settings;

Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;

}

function insertDefaultAgenda(event) {

    var subject = _settings.get("subject")
    console.log(subject + " - sub1")

    setTextToSubject(subject, event);

    //var body = _settings.get("body")
    //console.log(body + " - body7");

    //await setHTMLToBody(body, event);

   //event.completed();
}
    

function setTextToSubject(text, event) {

    Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, 
        function (asyncResult){
          if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
            statusUpdate(icon, "inserted successfully.");
          }
          else {
            Office.context.mailbox.item.notificationMessages.addAsync("addTextError", {
              type: "errorMessage",
              message: "Failed to insert " + asyncResult.error.message
            });
          }
          event.completed();
        });
}


function statusUpdate(icon, text) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
      type: "informationalMessage",
      icon: icon,
      message: text,
      persistent: false
    });
  }