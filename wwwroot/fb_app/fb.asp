<script language="javascript" runat="server" src="json2.asp"></script>
<!-- #INCLUDE FILE="fb_app.asp" -->
<%




'' A call to main gets the whole thing started

main

function main
	dim strJSON
	dim URL
	dim sToken
    dim user
    dim loc




        url = "http://localhost/fb_app/me.txt"
        strJSON = get_page_contents( URL ) 

        Set user = JSON.parse( strJSON )

        Set loc = JSON.parse( JSON.stringify( user.favorite_teams ) )

Response.write loc.name
 
        
    
end function

%>




<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml"
      xmlns:fb="http://www.facebook.com/2008/fbml">
  <body>


    <div id="fb-root"></div>
    <script src="http://connect.facebook.net/en_US/all.js"></script>
    <script>
      FB.init({appId: '<%= FACEBOOK_APP_ID %>', status: true,
               cookie: true, xfbml: true});
                FB.Event.subscribe('auth.login', function(response) {
                window.location.reload();
      });
    </script>
  </body>
</html>




