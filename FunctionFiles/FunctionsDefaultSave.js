
async function saveDefaultAgenda(event){

/*     _mailbox.item.body.getAsync(
        Office.CoercionType.Html, 
        function (bodyResult){
            _settings.set("body", bodyResult.value);
            _settings.saveAsync();        
        }
    ); */

    let result = await getBody();
    await setBody(result);

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


function setBody(body){

    return new Promise(function (resolve, reject) {
           
        try{
    
            _settings.set("body", body);
            _settings.saveAsync();

            resolve();
        }
        catch (error){
              reject();
        }
    })
}