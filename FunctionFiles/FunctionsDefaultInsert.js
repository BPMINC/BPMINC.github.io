
function insertDefaultAgenda(event) {

    var subject = _settings.get("subject")
    setTextToSubject(subject, event);

    var body = _settings.get("body")
    setHTMLToBody(body, event);
}

function setTextToSubject(text, event) {

    _mailbox.item.subject.setAsync(
        text,         
        function (asyncResult){
          statusUpdate(asyncResult);
          event.completed();
        });
}

function setHTMLToBody(html, event) {

    _mailbox.item.body.setSelectedDataAsync(
        html, 
        { coercionType: Office.CoercionType.Html }, 
        function (asyncResult){
          statusUpdate(asyncResult);
          event.completed();
        });
}