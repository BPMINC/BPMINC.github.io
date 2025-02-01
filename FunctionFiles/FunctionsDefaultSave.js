


Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;

}

function saveDefaultAgenda(event){

    _mailbox.item.body.getAsync(
        Office.CoercionType.Html, (bodyResult) => {
            _settings.set("body", bodyResult.value);
            _settings.saveAsync();

            event.completed()
        }
    );

}