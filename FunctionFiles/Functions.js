// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

Office.initialize = function () {
}

// Adds text into the body of the item, then reports the results
// to the info bar.
/* function addTextToBody(text, icon, event) {
  Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, 
    function (asyncResult){
      if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        statusUpdate(icon, "\"" + text + "\" inserted successfully.");
      }
      else {
        Office.context.mailbox.item.notificationMessages.addAsync("addTextError", {
          type: "errorMessage",
          message: "Failed to insert \"" + text + "\": " + asyncResult.error.message
        });
      }
      event.completed();
    });
} */

function addHTMLToBody(text, icon, event) {
  const mailItem = Office.context.mailbox.item;
  const base64String =
    "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAnUExURQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAN0S+bUAAAAMdFJOUwAQIDBAUI+fr7/P7yEupu8AAAAJcEhZcwAADsMAAA7DAcdvqGQAAAF8SURBVGhD7dfLdoMwDEVR6Cspzf9/b20QYOthS5Zn0Z2kVdY6O2WULrFYLBaLxd5ur4mDZD14b8ogWS/dtxV+dmx9ysA2QUj9TQRWv5D7HyKwuIW9n0vc8tkpHP0W4BOg3wQ8wtlvA+PC1e8Ao8Ld7wFjQtHvAiNC2e8DdqHqKwCrUPc1gE1AfRVgEXBfB+gF0lcCWoH2tYBOYPpqQCNwfT3QF9i+AegJfN8CtAWhbwJagtS3AbIg9o2AJMh9M5C+SVGBvx6zAfmT0r+Bv8JMwP4kyFPir+cswF5KL3WLv14zAFBCLf56Tw9cparFX4upgaJUtPhrOS1QlY5W+vWTXrGgBFB/b72ev3/0igUdQPppP/nfowfKUUEFcP207y/yxKmgAYQ+PywoAFOfCH3A2MdCFzD3kdADBvq10AGG+pXQBgb7pdAEhvuF0AIc/VtoAK7+JciAs38KIuDugyAC/v4hiMCE/i7IwLRBsh68N2WQjMVisVgs9i5bln8LGScNcCrONQAAAABJRU5ErkJggg==";

  // Get the current body of the message or appointment.
  mailItem.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
    if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
      // Insert the Base64-encoded image to the beginning of the body.
      const options = { isInline: true, asyncContext: bodyResult.value };
      mailItem.addFileAttachmentFromBase64Async(base64String, "sample.png", options, (attachResult) => {
        if (attachResult.status === Office.AsyncResultStatus.Succeeded) {
          let body = attachResult.asyncContext;
          body = body.replace("<p class=MsoNormal>", `<p class=MsoNormal><img src="cid:sample.png">`);

          mailItem.body.setAsync(body, { coercionType: Office.CoercionType.Html }, (setResult) => {
            if (setResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("Inline Base64-encoded image added to the body.");
            } else {
              console.log(setResult.error.message);
            }
          });
        } else {
          console.log(attachResult.error.message);
        }
      });
    } else {
      console.log(bodyResult.error.message);
    }
  });
}


function addDefaultMsgToBody(event) {
  addHTMLToBody("<b>This is the test agenda text</b>", "blue-icon-16", event);
}

/* function addDefaultMsgToBody(event) {
  addTextToBody("This is the R2R agenda text", "blue-icon-16", event);
} */

function addMsg1ToBody(event) {
  addTextToBody("This is the P2P agenda text", "red-icon-16", event);
}

// Gets the subject of the item and displays it in the info bar.
function getSubject(event) {
  var subject = Office.context.mailbox.item.subject;
  
  Office.context.mailbox.item.notificationMessages.addAsync("subject", {
    type: "informationalMessage",
    icon: "blue-icon-16",
    message: "Subject: " + subject,
    persistent: false
  });
  
  event.completed();
}

// Gets the item class of the item and displays it in the info bar.
function getItemClass(event) {
  var itemClass = Office.context.mailbox.item.itemClass;
  
  Office.context.mailbox.item.notificationMessages.addAsync("itemClass", {
    type: "informationalMessage",
    icon: "red-icon-16",
    message: "Item Class: " + itemClass,
    persistent: false
  });
  
  event.completed();
}

// Gets the date and time when the item was created and displays it in the info bar.
function getDateTimeCreated(event) {
  var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
  
  Office.context.mailbox.item.notificationMessages.addAsync("dateTimeCreated", {
    type: "informationalMessage",
    icon: "red-icon-16",
    message: "Created: " + dateTimeCreated.toLocaleString(),
    persistent: false
  });
  
  event.completed();
}

// Gets the ID of the item and displays it in the info bar.
function getItemID(event) {
  // Limited to 150 characters max in the info bar, so 
  // only grab the first 50 characters of the ID
  var itemID = Office.context.mailbox.item.itemId.substring(0, 50);
  
  Office.context.mailbox.item.notificationMessages.addAsync("itemID", {
    type: "informationalMessage",
    icon: "red-icon-16",
    message: "Item ID: " + itemID,
    persistent: false
  });
  
  event.completed();
}