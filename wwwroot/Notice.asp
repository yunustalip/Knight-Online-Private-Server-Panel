<%
response.charset="iso-8859-9"
response.expires=0
sesid=Session.Sessionid

noti=split(application("notice"),"|")

if application("notice")<>"" Then
if noti(0)<>"" and datediff("s",noti(1),now())<60 and instr(application("noticeuser"),sesid)=0 Then

Response.Write "<img src=""imgs/noticesol.gif"" align=""top"" style=""position:relative;top:0px""><marquee width=""90%"" style=""background-color:#000;"" height=""27"" valign=""bottom"">####NOTICE: "&noti(0)&" ####</marquee><img src=""imgs/noticesag.gif"" align=""top"" style=""position:relative;top:0px"">"
application("noticeuser")=application("noticeuser")+sesid+"|"
End If

End If
%>