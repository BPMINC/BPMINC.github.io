//Purpose: Defines an expression (not a definition) to be evaluated
//use variables outside the blocks for global access
// use let to override the variable values upon loading
let _mailbox
let _settings;

Office.initialize = function () {
  
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;

}