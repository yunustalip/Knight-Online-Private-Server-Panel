<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<%Response.expires=0 
If Not Request.ServerVariables("Script_Name")="/404.asp" Then
yn("/Download")
End If

REFERER_URL = Request.ServerVariables("HTTP_REFERER")
s=Request.ServerVariables("Script_Name")
If InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
Else
REFERER_DOMAIN = Left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If


If REFERER_DOMAIN="http://"&Request.ServerVariables("server_name") or  REFERER_DOMAIN="http://www."&Request.ServerVariables("server_name") Then
Else
yn("/Download")
End If

Dim MenuAyar,ksira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Download'")
If MenuAyar("PSt")=1 Then%>
<meta http-equiv="content-type" content="text/html; charset=windows-1254">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" >
<%Response.Charset = "iso-8859-9"
dim down
set down=Conne.Execute("select * from download")%>
<br /><img src="imgs/dosyalar.gif" /><br /><br /><br />
<table width="410" height="78" border="1" cellpadding="0" cellspacing="0">
  <% if not down.eof Then
	 do while not down.eof %>
  <tr align="center">
    <td bgcolor="#333333"><span style="font-weight: bold; color: #FFFFFF">Dosya Adı: <%=down("dosyaismi")%></span></td>
  </tr>
	<tr>
	  <td align="center" bgcolor="#e7dbc3"><br /><span style="font-weight: bold">Açıklama:</span> <%=down("aciklama")%><br /><br /></td>
	</tr>
	<tr><td align="left" bgcolor="#e7dbc3"><a href="<%=down("adres")%>">Download</a><br />
	<%=down("tarih")%></span></td>
  </tr>
  <%down.movenext
  loop
  End If %>
</table>
<%else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafından kapatılmıştır.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>