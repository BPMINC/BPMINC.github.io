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