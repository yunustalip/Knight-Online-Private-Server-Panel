<script language="javascript" runat="server" src="json2.asp"></script>
<!-- #INCLUDE FILE="fb_app.asp" -->
<%


''*** Logout Feature Block
'' The block of code below does not work correctly.
'' I have not been able to correctly implement a logout feature
'' If you have any ideas please share!
'' -- Larry
''
''
'' A feature to support logging out
''if request.querystring("cmd")="logout" then
''    response.cookies("fbs_" & FACEBOOK_APP_ID)=""
''    response.cookies("fbs_" & FACEBOOK_APP_ID).expires = Date()-10
''    response.redirect("fb.asp?cmd=login")
''end if    
''
''*** End Logout Feature Block




'' A call to main gets the whole thing started

main

function main
	dim strJSON
	dim URL
	dim sToken
    dim user
    dim loc


    set cookie = get_facebook_cookie( FACEBOOK_APP_ID, FACEBOOK_SECRET )

    if cookie.count > 0 then 
        response.write "Logged In!<br/>"
        
        '' Use the access token to get the userinfo
        '' as a JSON a string from Facebook
        sToken = cookie("access_token")
        url = "https://graph.facebook.com/me?access_token=" & sToken
        strJSON = get_page_contents( URL ) 

         '' Now that we have a json string from Facebook
         '' Use the json object from json2.asp to 
         '' convert the JSON content from Facebook into
         '' a user object so we can access properties 
         '' in the following fashion user.id
        set user = JSON.parse( strJSON )
        
        '' 
        '' Parse the hometown json, I couldn't figure out
        '' how to do this using user.hometown.name but I 
        '' suspect a get() call is involved (see json2 docs)
        '' 
        set loc = JSON.parse( JSON.stringify( user.hometown ) )

        ''on error resume next
        '' at this point you would add/update the user 
        '' record in your db 

        '' Here's how the data is accessed
        response.write user.id & "<br/>"
        response.write user.first_name & "<br/>"
        response.write user.last_name & "<br/>"
        response.write user.link & "<br/>"
        response.write "<img src='http://graph.facebook.com/" & user.id & "/picture'><br/>"
        response.write loc.name & "<br/>"
        response.write user.email & "<br/>"
        response.write cookie("access_token")
        response.write "<br/>"
 
        
    end if
    
end function

%>




<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml"
      xmlns:fb="http://www.facebook.com/2008/fbml">
  <body>
    <% if cookie.count > 0  then %>
      Your user ID is <%= cookie("uid") %>
    <%  else  %>
      <fb:login-button perms="email"></fb:login-button>
    <% end if %>

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




