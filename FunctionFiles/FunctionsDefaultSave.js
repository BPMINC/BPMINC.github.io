
async function saveDefaultAgenda(event){

    let subject = await getSubject();
    await setSubject(subject);

    let body = await getBody();
    await setBody(body);

}

function getSubject(){

    return new Promise(function (resolve, reject) {

        try {
            Office.context.mailbox.item.subject.getAsync(
                Office.CoercionType.Text, 
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(result.value);
                    } else {
                        reject(result.error);
                    }
                }
            );
        } catch {
            reject("Unable to get email subject text.");
        }
    })
}


function setSubject(text){

    return new Promise(function (resolve, reject) {
           
        try{
    
            _settings.set("subject", text);
            _settings.saveAsync();

            resolve();
        }
        catch (error){
              reject();
        }
    })
}


function getBody(){

    return new Promise(function (resolve, reject) {

        try {
            Office.context.mailbox.item.body.getAsync(
                Office.CoercionType.Html, 
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(result.value);
                    } else {
                        reject(result.error);
                    }
                }
            );
        } catch {
            reject("Unable to get email body text.");
        }
    })
}


function setBody(html){

    return new Promise(function (resolve, reject) {
           
        try{
    
            _settings.set("body", html);
            _settings.saveAsync();

            resolve();
        }
        catch (error){
              reject();
        }
    })
}