<script language="javascript" runat="server" src="json2.asp"></script>
<!-- #INCLUDE FILE="fb_app.asp" -->
<%Session.CodePage=1254
response.charset="windows-1254"




json_str = get_page_contents("http://handlers.karnaval.com/karnaval.functions.php?command=getActiveShow&radio_id=1")
set user = JSON.parse(Replace(Replace(json_str, "(",""),")","") )

	on error resume next



DjImage = user.image_url
ProgramName = user.show_name
DjLink = user.program.get(0).link
DjName = user.show_description
DjTime = Left(user.time_start,5) & " - " & Left(user.time_end,5)
Response.Write "<table><tr><td rowspan=""8"" valign=""top"" width=""90"">"&vbcrlf
If Len(djimage)>0 Then
Response.Write "<img src="""&djimage&""" style=""border: 3px solid rgb(221, 0, 0);"" alt=""djimage""/>"&vbcrlf
End If
Response.Write "</td></tr>"&vbcrlf
If Len(programname)>0 Then
Response.Write "<tr><td class=""Text""><strong>Program: </strong><a href="""&DjLink&""" target=""_blank"">"&ProgramName&"</a></td></tr>"&vbcrlf
End If
Response.Write "<tr><td class=""Text"">"&vbcrlf
If Len(DjLink)>0 Then
Response.Write "<div style=""float:left""><strong>DJ: </strong><a href="""&DjLink&""" target=""_blank"">"&DjName&"</a></div>"&vbcrlf
Else
Response.Write "<div style=""float:left""><strong>DJ: </strong>"& DjName &"</div>"&vbcrlf
End If
Response.Write "<div style=""float:left;color:#FF0000;font-weight:bold;text-decoration: blink;"">&nbsp;Yayında !</div>"&vbcrlf
Response.Write "<div style=""width:130px"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&Djtime&"</div>"&vbcrlf
%>
  <tr>
     <td align="center"><a href="http://www.superfm.com.tr/" target="_blank">Yayın Akışı</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://www.superfm.com.tr" target="_blank"><img src="images/frekans.gif" border="0" width="16" height="16" align="absmiddle" alt="frekans"/> Frekanslar</a></td>
  </tr>
  <tr>
    <td align="center" class="Text"><strong>Telefon: </strong>+90 212 368 6200</td>
  </tr>
  <tr>
    <td align="center" class="Text"><strong>E-posta: </strong>iletisim@superfm.com.tr</td>
  </tr>

<%
Response.Write"</table></td>"
Response.Write "  </tr>"



%>