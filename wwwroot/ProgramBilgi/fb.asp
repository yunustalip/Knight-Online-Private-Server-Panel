<% Option Explicit %>
<script language="javascript" runat="server" src="json2.asp"></script>
<!-- #INCLUDE FILE="fb_app.asp" -->

<%
'' JSON 2 Library from: 
''   https://github.com/nagaozen/asp-xtreme-evolution/tree/master/lib/axe/classes/Parsers
''

main

function main
	dim app_id
	dim app_secret
	dim my_url
	dim dialog_url
	dim token_url
	dim resp
	dim token
	dim expires
	dim graph_url
	dim json_str
	dim user
	dim code
	dim strLocation 
	dim strEducation
	dim strEmail
	dim strFirstName
	dim strLastName
	dim strID

	
	json_str = get_page_contents("http://www.karnaval.com/songs.php?radio=1")


	set user = JSON.parse( json_str )

	


	'' Handling properties that might not be there
	on error resume next
if err.number<>0 then
response.write err.description
end if


	   response.write user.currentsong.get(0).artist&" - "& user.currentsong.get(0).Title
    response.write "JSON String: <br/>"
   	response.write json_str
end function    


%>