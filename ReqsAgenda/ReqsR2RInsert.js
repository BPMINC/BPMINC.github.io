
async function insertReqsR2RAgenda(event) {

    var subject = getTextForSubject();
    await setTextToSubject(subject);
  
    var body = getHTMLForBody();
    await setHTMLToBody(body);
  
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

    return "Hi Team,<br /><br />Please join us to review the Record To Report \
    business requirements.<b><i>Meeting Objective</i></b><br/><br/>\
    The objective of this session is for our team to gather \
    a solid understanding of your AP processes from vendor \
    creation, vendor bills and associated approvals, vendor \
    payments and advanced electronic payments and expense reporting\
    <br/><br/><br/><b><i>Meeting Agenda</i></b><br/>\
    <ul><li>Vendor Master</li><li>Employee Master</li>\
    <li>Vendor Bills</li><li>Vendor Payments</li>\
    <li>Expense Reports</li><li>Fixed Assets</li></ul>\
    <br /><br />Thanks,<br />Joe"
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