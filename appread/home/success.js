jQuery(document).ready(function(){
  jQuery("#closeId").click(function (){
    window.localStorage.setItem("code", window.location.href);
    self.close();
  });
});