(function(){
  'use strict';

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();
      jQuery("#loginDiv").css("display", "none");
      jQuery("#loading").css("display", "block");

      //Get configuration data
      var client_id;
      var client_secret;
      var redirect_uri;
      var resource;
      var proxy;
      var realm;
      var site_url;
      var list_title;


      jQuery.getJSON( "./config.json", function( data ) {
         client_id = data.client_id;
         client_secret = data.client_secret;
         redirect_uri = data.redirect_uri;
         resource = data.resource;
         proxy = data.proxy;
         realm = data.realm;
         site_url = data.site_url;
         list_title = data.list_title;
 
         GetAccessToken();
      });
            
      
      
    jQuery("#btnAddSite").click(function (){

                var win = window.open("https://localhost:8443/appread/home/popup.html", "", "width=720, height=300, scrollbars=0, toolbar=0, menubar=0, resizable=0, status=0, titlebar=0");
                jQuery("#loginDiv").css("display", "none");
                jQuery("#loading").css("display", "block");
                if (window.focus) {
                  win.focus()}
                var winTimer = window.setInterval(function()
                {
                    if (win.closed !== false)
                    {
                      
                        // !== is required for compatibility with Opera
                        window.clearInterval(winTimer);
                        var popupUrl = window.localStorage.getItem("code");
                        var code = getUrlParameter(popupUrl, 'code');
                       
                        //Access code
                        var datas = 'grant_type=authorization_code&client_id='+client_id+'&client_secret='+client_secret+'&code='+code+'&redirect_uri='+redirect_uri+'&resource='+resource;
                        
                        jQuery.ajax({
                        url: proxy + "https://accounts.accesscontrol.windows.net/"+realm+"/tokens/OAuth/2",
                        type: "POST",
                        headers: { "Content-Type": "application/x-www-form-urlencoded"},
                        data: datas,
                        crossDomain: true,
                        contentType: "application/x-www-form-urlencoded",
                            success: function (data) {
                                window.localStorage.setItem("access_token", data.access_token);
                                window.localStorage.setItem("refresh_token", data.refresh_token);
                                jQuery("#loading").css("display", "none");
                                jQuery("#mainForm").css("display", "block");
                            },
                            error: function (data) {
                                
                            }
                        });
                            }
                        }, 200);
                  
            });
 
function GetAccessToken()
{
      var code = window.localStorage.getItem("refresh_token");
      //Refresh code
      var datas = 'grant_type=refresh_token&client_id='+client_id+'&client_secret='+client_secret+'&refresh_token='+code+'&redirect_uri='+redirect_uri+'&resource='+resource;
      
                jQuery.ajax({
                url: proxy + "https://accounts.accesscontrol.windows.net/"+realm+"/tokens/OAuth/2",
                type: "POST",
                headers: { "Content-Type": "application/x-www-form-urlencoded"},
                data: datas,
                crossDomain: true,
                contentType: "application/x-www-form-urlencoded",
                success: function (data) {
                    window.localStorage.setItem("access_token", data.access_token);
                    //window.localStorage.setItem("refresh_token", data.refresh_token);
                    jQuery("#loading").css("display", "none");
                    jQuery("#mainForm").css("display", "block");
                    var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
                    var itemId = item.subject.split('-');
                    getDocumentById(itemId[1]);
                },
                error: function (data) {
                    jQuery("#loginDiv").css("display", "block");
                    jQuery("#loading").css("display", "none");
                }
            });
}

      jQuery("#approve").click(function()
      {
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        var itemId = item.subject.split('-');
        var accessToken = window.localStorage.getItem('access_token');
        jQuery("#loading").css("display", "block");
        jQuery("#mainForm").css("display", "none");
        jQuery.ajax({
                url: site_url + "/_api/contextinfo",
                type: "POST",
                headers: { "Accept": "application/json;odata=verbose", 
              "Authorization": "Bearer " + accessToken,
              "Access-Control-Allow-Origin": "*"},
                success: function (data) {
                  var digital = data.d.GetContextWebInformation.FormDigestValue;
                  var itemProperties = {'DocStatus':'Approved'};
                  var itemPayload = {
                    '__metadata': {'type': getItemTypeForListName(list_title)}
                  };
                  for(var prop in itemProperties){
                        itemPayload[prop] = itemProperties[prop];
                  }
                  var body = JSON.stringify({ '__metadata': { 'type': 'SP.Data.DocumentsForApproveListItem' }, 'DocStatus': 'Approved'});
                  jQuery.ajax({
                        url: site_url + "/_api/web/lists/getbytitle('"+list_title+"')/items("+itemId[1]+")",
                        type: "POST",
                        data: JSON.stringify(itemPayload),
                        contentType: "application/json;odata=verbose",
                        headers: { "Accept": "application/json;odata=verbose", 
                        "X-RequestDigest": digital,
                        "X-HTTP-Method": "MERGE",
                        "IF-MATCH": "*",
                      "Authorization": "Bearer " + accessToken
                      
                      },
                        success: function (data) {
                            jQuery("#loading").css("display", "none");
                            jQuery("#mainForm").css("display", "none");
                            jQuery("#success").css("display", "block");
                        },
                        error: function (data) {
                            jQuery("#loading").css("display", "none");
                            jQuery("#mainForm").css("display", "block");
                            jQuery("#success").css("display", "none");
                        }
                    });
                   
                },
                error: function (data) {
                    //return null;
                }
            });

        
      });
      var site;
var getUrlParameter = function getUrlParameter(url, name) {
    var results = new RegExp('[\?&]' + name + '=([^&#]*)').exec(url);
	  return results[1] || 0;
};
function getItemTypeForListName(name) {
    return"SP.Data." + name.charAt(0).toUpperCase() + name.slice(1) + "ListItem";
}

function getDocumentById(itemId)
{
          var accessToken = window.localStorage.getItem('access_token');
          jQuery.ajax({
                url: site_url+ "/_api/web/lists/getbytitle('"+list_title+"')/items("+itemId+")",
                type: "GET",
                headers: { "Accept": "application/json;odata=verbose", 
              "Authorization": "Bearer " + accessToken,
              "Access-Control-Allow-Origin": "*"},
                success: function (data) {
                    document.getElementById("Title").innerHTML = data.d.Title;
                    document.getElementById("DocNotes").innerHTML = data.d.DocNotes;
                    var monthNames = [
                        "January", "February", "March",
                        "April", "May", "June", "July",
                        "August", "September", "October",
                        "November", "December"
                      ];      
                    var date = new Date(data.d.Created);
                    var day = date.getDate();
                    var monthIndex = date.getMonth();
                    var year = date.getFullYear();
                    document.getElementById("Created").innerHTML = day + " " + monthNames[monthIndex] + " " + year;
                },
                error: function (data) {
                    jQuery("#loginDiv").css("display", "block");
                    jQuery("#loading").css("display", "none");
                }
            });
}
    });
  };


  // Displays the "Subject" and "From" fields, based on the current mail item
  function displayItemDetails(){
    var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
    jQuery('#subject').text(item.subject);
   
    var from;
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      from = Office.cast.item.toMessageRead(item).from;
    } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      from = Office.cast.item.toAppointmentRead(item).organizer;
    }

    if (from) {
      jQuery('#from').text(from.displayName);
      jQuery('#from').click(function(){
        app.showNotification(from.displayName, from.emailAddress);
      });
    }
  }
})();
