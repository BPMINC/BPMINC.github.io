
async function insertDefaultAgenda(event) {

  var subject = _settings.get("subject")
  await setTextToSubject(subject);

  var body = _settings.get("body")
  await setHTMLToBody(body);

  event.completed();
}

function setTextToSubject(text) {

  return new Promise(function (resolve, reject) {
    try{

      _mailbox.item.subject.setAsync(
        text,         
        function (asyncResult){
          statusUpdate(asyncResult);
          resolve();
        }
      );
    }
    catch (error){
          reject();
    }
  })
}

function setHTMLToBody(html) {

  return new Promise(function (resolve, reject) {
    try{

      _mailbox.item.body.setSelectedDataAsync(
        html, 
        { coercionType: Office.CoercionType.Html }, 
        function (asyncResult){
          statusUpdate(asyncResult);
          resolve();
        }
      );
    }
    catch (error){
          reject();
    }
  })
}