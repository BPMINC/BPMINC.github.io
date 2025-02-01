
function saveDefaultAgenda(event){

    _mailbox.item.body.getAsync(
        Office.CoercionType.Html, (bodyResult) => {
            _settings.set("body", bodyResult.value);
            _settings.saveAsync();

            event.completed();
        }
    );

}