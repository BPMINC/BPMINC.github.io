

let _mailbox
let _settings;

Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;

  var subject = _settings.get("subject")
  $('#subjectToSave').val(subject);

  var body = _settings.get("body")
  $('#bodyToSave').val(body);

  $('#save').click(saveAgenda);
}
  
  function saveAgenda(){

    setSubject();
    setBody();  

    saveSettings();
    
  }

  function setSubject(){
      var text = $('#subjectToSave').val();
      _settings.set("subject", text);;
  }

  function setBody(){
      var html = $('#bodyToSave').val();
      _settings.set("body", html);       
  }

  function saveSettings(){

    _settings.saveAsync(function (asyncResult) {

      // Display the result to the user
      if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        _mailbox.item.notificationMessages.replaceAsync("status", {
          type: "informationalMessage",
          icon: "icon-16",
          message: "Save successful",
          persistent: false
        });        
      }
      else {
        _mailbox.item.notificationMessages.addAsync("error", {
          type: "errorMessage",
          message: "Save Failed - " + asyncResult.error.message,
          persistent: false
        }); 
      }
    })
  }