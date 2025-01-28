
async function insertDefaultAgenda(event) {

  var subject = Office.context.roamingSettings.get("subject");
  console.log(subject + " - sub")

  await setTextToSubject(subject, "icon-16");

  var body = Office.context.roamingSettings.get("body");
  console.log(body + " - body");

  await setHTMLToBody(body, "icon-16");
}



async function setTextToSubject(text, icon, event) {

  Office.context.mailbox.item.subject.setAsync(text, 

    function (asyncResult){
      if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
          type: "informationalMessage",
          icon: icon,
          message: "Success",
          persistent: false
        });
      }
      else {
        Office.context.mailbox.item.notificationMessages.addAsync("error", {
          type: "errorMessage",
          message: "Failed - " + asyncResult.error.message,
          persistent: false
        });
      }
      
      return new Promise((resolve, reject) => {  
        // Fake success  
        resolve("success");
      });

    }

  ); 
  event.completed();
} 

async function setHTMLToBody(html, icon, event) {
  Office.context.mailbox.item.body.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, 
    function (asyncResult){
      if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
          type: "informationalMessage",
          icon: icon,
          message: "Success",
          persistent: false
        });
      }
      else {
        Office.context.mailbox.item.notificationMessages.addAsync("error", {
          type: "errorMessage",
          message: "Failed - " + asyncResult.error.message,
          persistent: false
        });
      }

      return new Promise((resolve, reject) => {  
        // Fake success  
        resolve("success");
      });      
    }
  );
  event.completed();
} 