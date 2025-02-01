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