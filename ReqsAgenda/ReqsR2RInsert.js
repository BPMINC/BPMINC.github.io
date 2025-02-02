
function insertReqsR2RAgenda(event) {

  var file = "file:///C:/Users/JosephSmith/OneDrive - BPM/Desktop/Personal/Templates/thisfile.txt";

  var rawFile = new XMLHttpRequest();
  rawFile.open("GET", file, false);
  rawFile.onreadystatechange = function () {
    if(rawFile.readyState === 4)  {
      if(rawFile.status === 200 || rawFile.status == 0) {
        var allText = rawFile.responseText;
        console.log(allText);
       }
    }
  }
  rawFile.send(null);
  

    //var subject = getTextForSubject();
    //await setTextToSubject(subject);
  
    //var body = getHTMLForBody();
    //await setHTMLToBody(body);
  
    //event.completed();
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