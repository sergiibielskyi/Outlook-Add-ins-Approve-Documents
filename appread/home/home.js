(function(){
  'use strict';

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();
      jQuery("#loginDiv").css("display", "none");
      jQuery("#loading").css("display", "block");
      GetAccessToken();
      //jQuery.get('https://login.microsoftonline.com/common/oauth2/authorize?client_id=0b33a287-62dd-407e-bbb2-b9fc497ec39d&scope=openid+profile&response_type=id_token&redirect_uri=https://mysps365.sharepoint.com&nonce=2234345623456456',
      //{'crossDomain': true,
      //'dataType': 'jsonp'},
      // function(serverResponse){
      //  document.getElementById("website").innerHTML = "test";
     // });
        /*jQuery.getJSON('https://login.microsoftonline.com/common/oauth2/authorize?client_id=0b33a287-62dd-407e-bbb2-b9fc497ec39d&scope=openid+profile&response_type=id_token&redirect_uri=https://mysps365.sharepoint.com&nonce=2234345623456456',
          {
      tags: "mount rainier",
      tagmode: "any",
      format: "jsonp"
  })
  .done(function(data) {

  });*/
  //if ($(“#txtSiteUrl”).val().length >= 10) {

    
    jQuery("#btnAddSite").click(function (){
                //var url = "https://login.microsoftonline.com";
                //if (url.charAt(url.length) != '/')
                  //  url += '/';
                //build a redirect URI
                //var redirect = encodeURI("https://localhost:44367/Site/Add") + “&state=” + $(“#hdnUserID”).val() + “|” + encodeURI(url.toLowerCase());
                //url += "/common/oauth2/authorize?client_id=0b33a287-62dd-407e-bbb2-b9fc497ec39d&scope=openid+profile&response_type=code&redirect_uri=https://localhost:8443/appread/home/home.html&nonce=2234345623456456";
               // url += "https://localhost:8443/appread/home/home.html";
                //var url = "https://mysps365.sharepoint.com/_layouts/oauthauthorize.aspx?client_id=36fa0000-3a88-4ca5-82cd-aa7903aef2e1&scope=Web.Read&response_type=code&redirect_uri=https://localhost:8443/appread/home/home.html";
                
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
                        var datas = 'grant_type=authorization_code&client_id=cb9db5fb-6864-46f3-8bac-c030803fa4f7%40675871be-bb22-4c9d-86f0-954f9cbef0fa&client_secret=gw8Nzr4YhRG4NguuemLn7cf3WwBrFdDj%2FFgd%2BP%2BHn3Q=&code='+code+'&redirect_uri=https://localhost:8443/appread/home/success.html&resource=00000003-0000-0ff1-ce00-000000000000%2Fmysps365.sharepoint.com%40675871be-bb22-4c9d-86f0-954f9cbef0fa';

                       
                        /*var iframe = document.createElement("iframe");
  var uniqueString = datas;
  document.body.appendChild(iframe);
  iframe.style.display = "none";
  iframe.contentWindow.name = uniqueString;


  // construct a form with hidden inputs, targeting the iframe
  var form = document.createElement("form");
  form.target = uniqueString;
  form.action = "https://accounts.accesscontrol.windows.net/675871be-bb22-4c9d-86f0-954f9cbef0fa/tokens/OAuth/2";
  form.method = "POST";

  // repeat for each parameter
  var input = document.createElement("input");
  input.type = "hidden";
  //input.name = "Content-Type";
  input.value = uniqueString;
  form.appendChild(input);

  document.body.appendChild(form);
  form.submit();*/
                        
                var proxy = 'https://cors-anywhere.herokuapp.com/';

                jQuery.ajax({
                url: proxy + "https://accounts.accesscontrol.windows.net/675871be-bb22-4c9d-86f0-954f9cbef0fa/tokens/OAuth/2",
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


                
                //jQuery(win.document.body).html(response);
                //675871be-bb22-4c9d-86f0-954f9cbef0fa realm
                //https://accounts.accesscontrol.windows.net/675871be-bb22-4c9d-86f0-954f9cbef0fa/tokens/OAuth/2?grant_type=authorization_code&client_id=0b33a287-62dd-407e-bbb2-b9fc497ec39d&code=AQABAAAAAADRNYRQ3dhRSrm-4K-adpCJ8hu7gducw83bqJUOPD4nMXe9GBgU7ssX7VMgilsaCVRKU1QhzmnPPfQKuIvSs0jmG54tBRKjhUKkK3jU-XnizbidxJE4CkLDTNLBqESjkVOFvmWrzv8t7MJWd8tn0R7bNZLOOxxgS_8VRoQX5hrml-DNKZJyU4M541XYjkQf3FGtP4F-LeELwvKt-27pwxXlrlanGdyXnqtKazqAwa1QChFDaULS9CG2UU5EsY_24UBT9uxGru2fQxU9Kw3wS8_U_7rzSvNIwSCS4w8svnmdfHLhL78EoMN4t0QwkzXuItTumq317dXr5btrtdQOugiXB1oqjMYNb-5pQHChx3uJbdZMumQ62MLxYwDv4aHKVvPZFsMooRZJodytbOtEVbObpHUEu2j3lmUCdfyNvdTNnIoZRD3BgKsuNbGkVWd9aFJEt6vd0E6JTHNQKHNT-l8rGVF0wSskCtHgjt4WkU4serC3dcPFypmPXuHzpQv9mGCaNmoCoWauwqPdzProwTJZ5_JZkgJRSD6WbJ4Ob2nO-0t0dD1vnTRV0Yo_CjvNQCF2g7P1R5f1jswrn2ZDko7yjQmwKOn7y3yRTophlvvAmhahodll0h_ihM2FugiYbUZiMiSbyexLk7Z1hacFGghkv_G41zgOFZJ6TMVA50UN_HtsJe8hKoU8k7KoDp1u5ksHq8zkKD427afAD0nnZYoQIAA&session_state=6c098cb8-a026-41bd-80db-0c885c8693ed&redirect_uri=https://localhost:8443/appread/home/home.html&resource=00000003-0000-0ff1-ce00-000000000000%2Fyour_site_name.sharepoint.com%40d2076ad6-6179-41cb-b792-24716a55ea90
                
                //client_id = 4e4a4b2e-b1a7-4a5a-a09f-ceb96393e386
                //client secret = k8i12mO52ympD2/ZUAHfe4W5tdbcR1luGs8wStqIc1Y=
                //realm = 675871be-bb22-4c9d-86f0-954f9cbef0fa
                //redirect = https://localhost:8443/appread/home/home.html
                //https://mysps365.sharepoint.com/_layouts/oauthauthorize.aspx?client_id=36fa0000-3a88-4ca5-82cd-aa7903aef2e1&scope=Web.Read&response_type=code&redirect_uri=https://localhost:8443/appread/home/home.html
                //code = IAAAACZPWVLA_A2kBQydgetnUq4123lRv19VFiN2EejWLcNk_jTRAGCtbkZ2X-R891cDS4ZC7s5C3J0wzkM7kvjiV4W9Us892YYwyXVHn1b1tAUUojFJey9n1uLt26vbwFczyWn0YeyS5V4MG3ZIj906VmzSicDApVj1dgCU01soIJ-2aAgCbI1Jv8xInDnU5efXaBN5I9OujyzezU0mf4zRUJnmYpJT3Mz5IhneWjTcbUSlC4yJARVw818d9DGabxFQnQMBdsecALdekzTADiBFyhz0sMHQ7a-h-2dvJpMyF7yKCWBW_Gu41wt4zddcHnHp9sMReugELnhCN2TVLunwWVDK2SVglstGAB6kyUTqG7cAbuDv7I0Hx4EuczaCPyYS0A
                  
            });
            //grant_type=authorization_code&client_id=36fa0000-3a88-4ca5-82cd-aa7903aef2e1%40675871be-bb22-4c9d-86f0-954f9cbef0fa&client_secret=YugAMCEm0Mosf2JIOX7df5rs0jmh82T6/No1fvkY3IY=&code=IAAAACZPWVLA_A2kBQydgetnUq4123lRv19VFiN2EejWLcNk_jTRAGCtbkZ2X-R891cDS4ZC7s5C3J0wzkM7kvjiV4W9Us892YYwyXVHn1b1tAUUojFJey9n1uLt26vbwFczyWn0YeyS5V4MG3ZIj906VmzSicDApVj1dgCU01soIJ-2aAgCbI1Jv8xInDnU5efXaBN5I9OujyzezU0mf4zRUJnmYpJT3Mz5IhneWjTcbUSlC4yJARVw818d9DGabxFQnQMBdsecALdekzTADiBFyhz0sMHQ7a-h-2dvJpMyF7yKCWBW_Gu41wt4zddcHnHp9sMReugELnhCN2TVLunwWVDK2SVglstGAB6kyUTqG7cAbuDv7I0Hx4EuczaCPyYS0A&redirect_uri=https://localhost:8443/appread/home/home.html&resource=00000003-0000-0ff1-ce00-000000000000%2Fmysps365.sharepoint.com%40675871be-bb22-4c9d-86f0-954f9cbef0fa
                //jQuery("#refreshModal").modal("show");
                //});
                //https://localhost:8443/appread/home/home.html?code=AQABAAAAAADRNYRQ3dhRSrm-4K-adpCJ8hu7gducw83bqJUOPD4nMXe9GBgU7ssX7VMgilsaCVRKU1QhzmnPPfQKuIvSs0jmG54tBRKjhUKkK3jU-XnizbidxJE4CkLDTNLBqESjkVOFvmWrzv8t7MJWd8tn0R7bNZLOOxxgS_8VRoQX5hrml-DNKZJyU4M541XYjkQf3FGtP4F-LeELwvKt-27pwxXlrlanGdyXnqtKazqAwa1QChFDaULS9CG2UU5EsY_24UBT9uxGru2fQxU9Kw3wS8_U_7rzSvNIwSCS4w8svnmdfHLhL78EoMN4t0QwkzXuItTumq317dXr5btrtdQOugiXB1oqjMYNb-5pQHChx3uJbdZMumQ62MLxYwDv4aHKVvPZFsMooRZJodytbOtEVbObpHUEu2j3lmUCdfyNvdTNnIoZRD3BgKsuNbGkVWd9aFJEt6vd0E6JTHNQKHNT-l8rGVF0wSskCtHgjt4WkU4serC3dcPFypmPXuHzpQv9mGCaNmoCoWauwqPdzProwTJZ5_JZkgJRSD6WbJ4Ob2nO-0t0dD1vnTRV0Yo_CjvNQCF2g7P1R5f1jswrn2ZDko7yjQmwKOn7y3yRTophlvvAmhahodll0h_ihM2FugiYbUZiMiSbyexLk7Z1hacFGghkv_G41zgOFZJ6TMVA50UN_HtsJe8hKoU8k7KoDp1u5ksHq8zkKD427afAD0nnZYoQIAA&session_state=6c098cb8-a026-41bd-80db-0c885c8693ed    
   // })
function GetAccessToken()
      {
                        var code = window.localStorage.getItem("refresh_token");
                        //Refresh code
                        var datas = 'grant_type=refresh_token&client_id=cb9db5fb-6864-46f3-8bac-c030803fa4f7%40675871be-bb22-4c9d-86f0-954f9cbef0fa&client_secret=gw8Nzr4YhRG4NguuemLn7cf3WwBrFdDj%2FFgd%2BP%2BHn3Q=&refresh_token='+code+'&redirect_uri=https://localhost:8443/appread/home/success.html&resource=00000003-0000-0ff1-ce00-000000000000%2Fmysps365.sharepoint.com%40675871be-bb22-4c9d-86f0-954f9cbef0fa';
      
    var proxy = 'https://cors-anywhere.herokuapp.com/';

                jQuery.ajax({
                url: proxy + "https://accounts.accesscontrol.windows.net/675871be-bb22-4c9d-86f0-954f9cbef0fa/tokens/OAuth/2",
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
                },
                error: function (data) {
                    jQuery("#loginDiv").css("display", "block");
                    jQuery("#loading").css("display", "none");
                }
            });
                  
  }

      jQuery("#test").click(function()
      {
        var accessToken = window.localStorage.getItem('access_token');
          jQuery.ajax({
                url: "https://mysps365.sharepoint.com/_api/web/lists/getbytitle('ManagersList')/items",
                type: "GET",
                headers: { "Accept": "application/json;odata=verbose", 
              "Authorization": "Bearer " + accessToken,
              "Access-Control-Allow-Origin": "*"},
                success: function (data) {
                    //alert('Successfully obtained data.');
                },
                error: function (data) {
                    //alert(data);
                }
            });
      });
      //getwebsite();
    });
  };
var site;
var getUrlParameter = function getUrlParameter(url, name) {
    var results = new RegExp('[\?&]' + name + '=([^&#]*)').exec(url);
	  return results[1] || 0;
};
function getwebsite()
{
    var clientContext = new SP.ClientContext('https://mysps365.sharepoint.com');
    this.oWebsite = clientContext.get_web();

    clientContext.load(this.oWebsite);

    clientContext.executeQueryAsync(
        Function.createDelegate(this, this.onQuerySucceeded), 
        Function.createDelegate(this, this.onQueryFailed)
    );  
}
function onQuerySucceeded(sender, args) {
    site = this.oWebsite.get_title();
}
    
function onQueryFailed(sender, args) {
    site = args.get_message();
}
  // Displays the "Subject" and "From" fields, based on the current mail item
  function displayItemDetails(){
    var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
    jQuery('#subject').text(item.subject);
    getwebsite();
    jQuery('#website').text(site);

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
