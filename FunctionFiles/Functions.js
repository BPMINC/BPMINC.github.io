
let _mailbox
let _settings;

Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;

}

function insertDefaultAgenda(event) {

    var subject = _settings.get("subject")
    console.log(subject + " - sub1")

    setTextToSubject(subject);

    //var body = _settings.get("body")
    //console.log(body + " - body7");

    //await setHTMLToBody(body, event);

    event.completed();
}
    

function setTextToSubject(text) {

    _mailbox.item.subject.setAsync(text, 
        
        function (asyncResult) {
            // Display the result to the user
            console.log("attempt result")
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                console.log(asyncResult)
              _mailbox.item.notificationMessages.replaceAsync("status", {
                type: "informationalMessage",
                icon: "icon-16",
                message: "Insert successful",
                persistent: true
              });        
            }
            else {
                console.log(asyncResult)
              _mailbox.item.notificationMessages.replaceAsync("error", {
                type: "errorMessage",
                message: "Insert Failed - " + asyncResult.error.message,
                persistent: true
              }); 
            }
            console.log("finished check")
        }
    ); 

    console.log("finish function")
}