# add-ins
Office Add-ins development
Outlook add-ins can help approve documents from SharePoint online via outlook client.
Fist you should configure config file parameters:
  1. client_id - this parameter contains two values (clientId@realm) that one of them you can get from SharePoint _layouts/15/AppInv.aspx and realm you can get from web request header (GET) _vti_bin/client.svc WWW-Authenticate →Bearer realm=... Parameter should be used in encoded format. 
  2. client_secret - this parameter you can get from SharePoint. Parameter should be used in encoded format. 
  3. redirect_uri - this parameter you can get from SharePoint _layouts/15/AppInv.aspx
  4. resource - this parameter contains three parameters: 1. audience orincipal ID, that you can get it from request to _vti_bin/client.svc and look to header response. 2. dns name of your SharePoint site. 3. realm.
  5. proxy - https://cors-anywhere.herokuapp.com/ but you can use any proxy just for test and debug using localhost. When you publish this solution to the web proxy is not required option. You should use proxy when you develop due to CORS https://developer.mozilla.org/en-US/docs/Web/HTTP/Access_control_CORS
  6. realm - this is your SharePoint realm
  7. site_url - your SharePoint web site name
  8. list_title - your SharePoint list title
  9. client_id_short - your app client id
  
  How it works
  User should recieve email from SharePoint workflow to approve document. Subject of email has to have list item id of document using this format - e.g. "Some document - ID". After user should click to the green button on the right side to activate Add-ins. The next step is to login to the site and click to approve the document. Access token is saved on local storage from browser and user will login automatically. Also it is possible to use outlook client to execute the same scenario.
  Note. SharePoint column calls DocStatus that Add-in will update. Type of this column should to be text.
