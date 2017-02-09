jQuery(document).ready(function(){
    //Get configuration data
      var client_id_short;
      var redirect_uri;
      var site_url;
      
      jQuery.getJSON( "./config.json", function( data ) {
         client_id_short = data.client_id_short;
         redirect_uri = data.redirect_uri;
         site_url = data.site_url;
         
         var url = site_url + "/_layouts/oauthauthorize.aspx?IsDlg=1&client_id="+client_id_short+"&scope=Web.Write&response_type=code&redirect_uri="+redirect_uri;
         window.location.href = url;
      });
    
    
});