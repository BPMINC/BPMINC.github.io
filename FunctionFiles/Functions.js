
let _mailbox
let _settings;

Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;

}

function insertDefaultAgenda(event){

    _mailbox.item.body.getAsync(
        Office.CoercionType.Html, (bodyResult) => {
            _settings.set("body", html);
            _settings.saveAsync();
        }
    );

    event.completed();

}

function tempfunc(event) {

    var subject = _settings.get("subject")
    setTextToSubject(subject, event);

    var body = _settings.get("body")
    setHTMLToBody(body, event);
}

function setTextToSubject(text, event) {

    _mailbox.item.subject.setAsync(
        text,         
        function (asyncResult){
          statusUpdate(asyncResult);
          event.completed();
        });
}

function setHTMLToBody(html, event) {

    _mailbox.item.body.setSelectedDataAsync(
        html, 
        { coercionType: Office.CoercionType.Html }, 
        function (asyncResult){
          statusUpdate(asyncResult);
          event.completed();
        });
}

function statusUpdate(asyncResult){
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        _mailbox.item.notificationMessages.replaceAsync("status", {
            type: "informationalMessage",
            icon: "icon-16",
            message: "Insert Successful",
            persistent: false
        });
    }
    else {
        _mailbox.item.notificationMessages.replaceAsync("error", {
        type: "errorMessage",
        message: "Save Failed - " + asyncResult.error.message
        });
    }
}