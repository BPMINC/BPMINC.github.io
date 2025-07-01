let _mailbox
let _settings;

Office.initialize = function () {

  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;

  const { fromEvent } = fileSelector;

  document.addEventListener('drop', async evt => {
    const files = await fromEvent(evt);

    console.log("start")
    console.log(files);
    console.log("finish")
  });

}