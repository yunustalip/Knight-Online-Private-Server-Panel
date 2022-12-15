<script language="javascript" runat="server" src="json2.asp"></script>
<!-- #INCLUDE FILE="fb_app.asp" -->
<%Session.CodePage=65001
response.charset="utf-8"




	json_str = get_page_contents("http://www.karnaval.com/songs.php?radio=4")
	set user = JSON.parse( json_str )

	on error resume next



DjImage = user.program.get(0).image
ProgramName = user.program.get(0).showName
DjLink = user.program.get(0).link
DjName = user.program.get(0).showDescription
DjTime = user.program.get(0).showTime

Response.Write "<table ><tr><td rowspan=""8"" valign=""top"" width=""90"">"
If Len(djimage)>0 Then
Response.Write "<img src="""&djimage&""" style=""border: 3px solid rgb(221, 0, 0);"">"
End If
Response.Write "</td></tr>"
If Len(programname)>0 Then
Response.Write "<tr><td class=""Text""><strong>Program: </strong><a href="""&DjLink&""" target=""_blank"">"&ProgramName&"</a></td></tr>"
End If
Response.Write "<tr><td class=""Text""><strong>DJ: </strong>"
If Len(DjLink)>0 Then
Response.Write "<a href="""&DjLink&""" target=""_blank"">"&DjName&"</a>"
Else
Response.Write DjName
End If
Response.Write "<blink><font color=""red""><b> Yayında !</b></font></blink><br>&nbsp;&nbsp;&nbsp;&nbsp;"&Djtime&""
%>
  <tr>
     <td align="center"><a href="http://www.joyturk.com.tr" target="_blank">Yayın Akışı</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://www.joyturk.com.tr" target="_blank"><img src="http://localhost/radyo/images/frekans.gif" border="0" width="16" height="16" align="absmiddle"> Frekanslar</a></td>
  </tr>
  <tr>
    <td align="center" class="Text"><strong>Telefon: </strong>+90 212 368 6200</td>
</tr>
<tr>
    <td align="center" class="Text"><strong>E-Posta: </strong>iletisim@joyturk.com.tr</td>
  </tr>

<%
Response.Write"</table></td>"
Response.Write "  </tr>"



%>