function statusUpdate(asyncResult, text){
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        _mailbox.item.notificationMessages.replaceAsync("status", {
            type: "informationalMessage",
            icon: "icon-16",
            message: text + " Successful",
            persistent: false
        });
    }
    else {
        _mailbox.item.notificationMessages.replaceAsync("error", {
        type: "errorMessage",
        message: text + " Failed - " + asyncResult.error.message
        });
    }
}