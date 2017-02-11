# add-ins
Office Add-ins development
Outlook add-ins can help approve documents from SharePoint online via outlook client.
Fist you should configure config file parameters:
  1. client_id - this parameter contains two values (clientId@realm) that one of them you can get from SharePoint _layouts/15/AppInv.aspx and realm you can get from web request header (GET) _vti_bin/client.svc WWW-Authenticate →Bearer realm=... Parameter should be used in encoded format. 
  2. client_secret - this parameter you can get from SharePoint. Parameter should be used in encoded format. 
  3. redirect_uri - this parameter you can get from SharePoint _layouts/15/AppInv.aspx
  4. resource - this parameter contains 
  5. proxy - 
  6. realm - 
  7. site_url - 
  8. list_title - 
  9. client_id_short - 
