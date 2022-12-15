<!--#include file="../_inc/conn.asp"-->
<%response.expires=0
Response.Charset = "iso-8859-9"
Dim MenuAyar,ksira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Chat'")
If MenuAyar("PSt")=1 Then
%>
<meta http-equiv="content-type" content="text/html; charset=windows-1254">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" >
<script language="JavaScript">
<!--
function launchChat() {
	var winl = (screen.width - 700) / 2;
	var wint = (screen.height - 500) / 2;
	window.open('chat/default.asp', 'ChatWindow', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=725,height=500,top='+wint+',left='+winl)
}

-->
</script>
<br><img src="../imgs/chat.gif" border="0">
<br><br><br>
<img src="chat/chatroom_image.asp">
<br>
<br>
<br>
<a href="javascript:launchChat();" class="link1"><img src="../imgs/chatrommlogin.gif" border="0"></a>
<% else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafından kapatılmıştır.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>