
async function saveDefaultAgenda(event){

    await saveSubject();
    await saveBody();

    event.completed();
}

function saveSubject(){

    return new Promise(function (resolve, reject) {

        try {
            _mailbox.item.subject.getAsync(                
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
  
                        _settings.set("subject", result.value);
                        _settings.saveAsync();
            
                        resolve();
                        
                    } 
                    else {
                        reject();
                    }            
                }
            );
        } 
        catch {
            reject("Unable to get email subject text");
        }
    })
}

function saveBody(){

    return new Promise(function (resolve, reject) {

        try {
            Office.context.mailbox.item.body.getAsync(
                Office.CoercionType.Html, 
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
  
                        _settings.set("body", result.value);
                        _settings.saveAsync();
            
                        resolve(result.value);
                        
                    } 
                    else {
                        reject(result.error);
                    }            
                }
            );
        } 
        catch {
            reject("Unable to get email body text.");
        }
    })
}