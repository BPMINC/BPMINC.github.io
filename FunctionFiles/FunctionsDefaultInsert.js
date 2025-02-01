
async function insertDefaultAgenda(event) {

  var subject = _settings.get("subject")
  await setTextToSubject(subject, event);

  var body = _settings.get("body")
  setHTMLToBody(body, event);
}

function setTextToSubject(text, event) {

  return new OfficeExtension.Promise(function (resolve, reject) {
    try{

      _mailbox.item.subject.setAsync(
        text,         
        function (asyncResult){
          statusUpdate(asyncResult);
          event.completed();
          resolve();
        }
      );
    }
    catch (error){
          reject();
    }
  })
}

function setHTMLToBody(html, event) {

  _mailbox.item.body.setSelectedDataAsync(
    html, 
    { coercionType: Office.CoercionType.Html }, 
    function (asyncResult){
      statusUpdate(asyncResult);
      event.completed();
    }
  );
}