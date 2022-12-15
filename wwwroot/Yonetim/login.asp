<link href="css/styles.css" rel="stylesheet" type="text/css" >
<br><br><br><br><br><br><br><br><br><br><br><center>
<!--#include file="../_inc/conn.asp"--><%
If Request.Querystring("islem")="login" Then%>
<!--#include file="../function.asp"-->
<!--#include file="../guvenlik.asp"-->
<%
Dim strAccountID,strPasswd,SecurityCode,LoginType
strAccountID = secur(Request.Form("strAccountID"))
strPasswd = secur(Request.Form("strPasswd"))
SecurityCode = secur(Request.Form("SecurityCode"))
LoginType = secur(Request.Form("logintype"))

If strAccountID="" Or strPasswd="" Then %>
<meta http-equiv="refresh" content="1;url=default.asp">
<table>
<tr><td align="center">

<img src="../imgs/18-1.gif">
<br /><br />
Kullanýcý Adý Veya Þifrenizi Kontrol Ediniz.
</td>
</tr>
</table>
<%Response.End

ElseIf SecurityCode<>Session("human-control_" & Session.SessionID) Then%>
<meta http-equiv="refresh" content="1;url=Default.asp">
<table>
<tr><td align="center"><img src="../imgs/18-1.gif">
<br /><br />Güvenlik Kodu Yanlýþ</td>
</tr>
</table>
<%Response.End

Else %>
</head>
<div align="center" class="style1">
<%Set userlogin = Conne.Execute("SELECT * FROM yonetim WHERE users = '"&strAccountID&"' AND pass = '"&strPasswd&"'")

If Not Userlogin.Eof Then


Session("strAccountID")=strAccountID
Session("strPasswd")=strPasswd
Session("durum")="esp"

%><font face="Verdana" style="font-size:9pt;"><b>Giriþ Baþarýlý</b></font><br /><br />
<font face="Verdana" style="font-size:9pt;"><i>Siteye Yönlendiriliyorsunuz..</i> </font><br />
<img src="../imgs/18-1.gif" />
<meta http-equiv="refresh" content="1;url=default.asp">
<%Response.End
Else %>
<meta http-equiv="refresh" content="1;url=default.asp">
<font face="Verdana" style="font-size:9pt;"><i>Kullanýcý Adý veya Parola Yanlýþ..</i></font>
<img src="../imgs/18-1.gif" /></div> 
<%Response.End
 End If 
End If 


Else

If Session("strAccountID")="" or Session("durum")="" Then
%>
  
<script type="text/javascript" language="JavaScript">
function reloadImage()
{
   document.images["simage"].src =  '../securityImage.asp?rand='+Math.random() * 1000000
}
</script>
Kullanýcý Adý ve þifrenizle giriþ yapýnýz. Eðer yönetici deðilseniz lütfen <a href="../default.asp">týklayýn</a><br />
<br /><form action="Login.asp?islem=login" method="POST" class="style1">
<table>
<tr><td colspan="2" align="center">Yönetim Sayfasý Administrator Giriþi</td></tr>
<tr><td>Kullanýcý Adý</td><td><input type="text" maxlength="12" name="strAccountID" value="<%=strAccountID%>"/></td></tr>
<tr><td>Þifre</td><td><input type="password"maxlength="12" name="strPasswd" /></td></tr>
<tr><td>Güvelik Kodu</td><td><input type="text" name="SecurityCode" id="SecurityCode" maxlength="8"> </td>
<td><img id="simage" src="../securityImage.asp">
<a href="#" onClick="reloadImage();return false;" style="position:relative;top:-11px">
<label for="SecurityCode">( Kodu Yenile )</label></a></td></tr>
<tr><td colspan="2"><input type="submit" value="Giriþ" style="width:300px"/></td></tr>
</table>
</form>
<%Else
Response.Redirect("Default.asp")
End If

End If %>