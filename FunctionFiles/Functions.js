
let _mailbox
let _settings;

Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;

}

function insertDefaultAgenda(event) {

    var subject = _settings.get("subject")
    console.log(subject + " - sub1")

    setTextToSubject(subject);

    //var body = _settings.get("body")
    //console.log(body + " - body7");

    //await setHTMLToBody(body, event);

    event.completed();
}
    

function setTextToSubject(text) {

    _mailbox.item.subject.setSelectedDataAsync(
        text,
        { coercionType: Office.CoercionType.Html },
        function (asyncResult) {
            myLogger("finished check");
        }
    ); 

    console.log("finish function")
}


function myLogger(message){
    return function(){
        console.log(message);
    }      
}