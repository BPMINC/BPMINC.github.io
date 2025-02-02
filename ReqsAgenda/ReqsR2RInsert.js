
function insertReqsR2RAgenda(event) {

  const url = "https://bpmcpa.app.box.com/file/1666928015557?s=7ajotsruy10tzhr952euf8efhp4nrp8r";
  const params = {method: "GET", mode: "cors"}
  }
  try {
    const response = fetch(url, params);
    if (!response.ok) {
      throw new Error(`Response status: ${response.status}`);
    }

    //const json = await response.json();
    console.log(response);

  } catch (error) {
    console.error("my error - " + error.message);
  }
  
  

    //var subject = getTextForSubject();
    //await setTextToSubject(subject);
  
    //var body = getHTMLForBody();
    //await setHTMLToBody(body);
  
    event.completed();
  }

  function getTextForSubject(){

    return "CLIENTNAME - R2R Requirements Gathering"

  }
  
  function setTextToSubject(text) {
  
    return new Promise(function (resolve, reject) {
      try{
  
        _mailbox.item.subject.setAsync(
          text,         
          function (asyncResult){
            statusUpdate(asyncResult,"Insert");
            resolve();
          }
        );
      }
      catch (error){
            reject();
      }
    })
  }

  function getHTMLForBody(){

    return "<b><i>Meeting Objective</i></b><br/><br/>\
    The objective of this session is for our team to gather \
    a solid understanding of your AP processes from vendor \
    creation, vendor bills and associated approvals, vendor \
    payments and advanced electronic payments and expense reporting\
    <br/><br/><br/><b><i>Meeting Agenda</i></b><br/>\
    <ul><li>Vendor Master</li><li>Employee Master</li>\
    <li>Vendor Bills</li><li>Vendor Payments</li>\
    <li>Expense Reports</li><li>Fixed Assets</li></ul>\
    "
  }
  
  function setHTMLToBody(html) {
  
    return new Promise(function (resolve, reject) {
      try{
  
        _mailbox.item.body.setSelectedDataAsync(
          html, 
          { coercionType: Office.CoercionType.Html }, 
          function (asyncResult){
            statusUpdate(asyncResult,"Insert");
            resolve();
          }
        );
      }
      catch (error){
            reject();
      }
    })
  }