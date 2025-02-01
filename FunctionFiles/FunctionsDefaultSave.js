
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

    event.completed();
}

function getBody(){

    return new Promise(function (resolve, reject) {
    
        try{
            let body;

            _mailbox.item.body.getAsync(
                Office.CoercionType.Html,
                function(asyncResult){
                    body = asyncResult;
                }
            );
            
            console.log("done " + body);
            resolve(body);
        }
        catch (error){
            reject();
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