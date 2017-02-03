(function(){
  'use strict';

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();

      jQuery('#set-subject').click(setSubject);
      jQuery('#get-subject').click(getSubject);
      jQuery('#add-to-recipients').click(addToRecipients);

      
    });
  };

  function setSubject(){
    //Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.setAsync('Hello world!');
    Office.context.mailbox.item.body.setAsync(
         "<b>(replaces all body, including threads you are replying to that may be on the bottom)</b>",
        { coercionType:"html", asyncContext:"This is passed to the callback" },
          function callback(result) {
          // Process the result
        });
  }
  

  function getSubject(){
    jQuery.$.get('https://login.microsoftonline.com/common/oauth2/authorize?client_id=0b33a287-62dd-407e-bbb2-b9fc497ec39d&scope=openid+profile&response_type=id_token&redirect_uri=https://mysps365.sharepoint.com&nonce=2234345623456456',
       function(serverResponse){
        document.getElementById("website").innerHTML = "test";
      });
         

   // })
      
      //getwebsite();
    
    Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.getAsync(function(result){
      app.showNotification('The current subject is', result.value);
    });
  }

  function addToRecipients(){
    var item = Office.context.mailbox.item;
    var addressToAdd = {
      displayName: Office.context.mailbox.userProfile.displayName,
      emailAddress: Office.context.mailbox.userProfile.emailAddress
    };

    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      Office.cast.item.toMessageCompose(item).to.addAsync([addressToAdd]);
    } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      Office.cast.item.toAppointmentCompose(item).requiredAttendees.addAsync([addressToAdd]);
    }
  }

})();
