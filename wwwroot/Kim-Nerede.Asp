<!--#include file="Function.asp"-->
<%If Not Request.ServerVariables("Script_Name")="/404.asp" Then
yn("/Kim-Nerede")
End If

REFERER_URL = Request.ServerVariables("HTTP_REFERER")

If InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
Else
REFERER_DOMAIN = Left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If


If REFERER_DOMAIN="http://"&Request.ServerVariables("server_name") or  REFERER_DOMAIN="http://www."&Request.ServerVariables("server_name") Then
Else
yn("/Kim-Nerede")
End If%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<body bgcolor="#F9EED8"><br><img src="imgs/kimnerede.gif"><br><br>
<iframe src="zone.asp" scrolling="no" frameborder="0"  allowtransparency="true" width="700" height="530" style="background:transparent;width:700;height:530;padding-left:15;padding-top:10"></iframe>
</body>