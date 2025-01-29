

let _mailbox;
let _settings;

Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
}

async function insertDefaultAgenda(event) {

    var subject = _settings.get("subject");
    console.log(subject + " - sub3")

    await setTextToSubject(subject, event);

    var body = _settings.get("body");
    console.log(body + " - body3");

    await setHTMLToBody(body, event);

    event.completed();
}
  
  
  
function setTextToSubject(text, event) {

    _mailbox.item.subject.setAsync(text, 

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

} 