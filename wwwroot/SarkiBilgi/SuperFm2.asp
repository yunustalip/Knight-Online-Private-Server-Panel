<%response.charset="utf-8"
Session.codepage=65001
%>
<script language="javascript" runat="server" src="json2.asp"></script>
<!-- #INCLUDE FILE="fb_app.asp" -->

<%

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

	on error resume next

   response.write "<Table><tr><td rowspan=""3""></td>"& user.currentsong.get(0).artist&" - "& user.currentsong.get(0).Title
   response.write "JSON String: <br/>"
   	response.write json_str
end function    


%>json_str = get_page_contents("http://www.karnaval.com/songs.php?radio=1")
	set user = JSON.parse( json_str )

	on error resume next

   response.write "<div id=""songduration"" style=""width:200px;border:solid #000 1px""><span style=""z-index:-1;"">123</span><div id=""songcurrent"" style=""background-color:#000;width:"& cint((user.currentsong.get(0).position/user.currentsong.get(0).duration)*200)&"px"">&nbsp;</div></div>"
   response.write "JSON String: <br/>"
   response.write json_str


%>
<script>

<%
response.write "var current="& cint(user.currentsong.get(0).position)&";"&vbcrlf
Response.Write "var duration="& cint(user.currentsong.get(0).duration)&";"&vbcrlf
Response.Write "document.getElementById('songcurrent').style.width=current/duration*200+'px'"%>



</script>