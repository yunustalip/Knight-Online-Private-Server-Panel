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
Kullan�c� Ad� Veya �ifrenizi Kontrol Ediniz.
</td>
</tr>
</table>
<%Response.End

ElseIf SecurityCode<>Session("human-control_" & Session.SessionID) Then%>
<meta http-equiv="refresh" content="1;url=Default.asp">
<table>
<tr><td align="center"><img src="../imgs/18-1.gif">
<br /><br />G�venlik Kodu Yanl��</td>
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

%><font face="Verdana" style="font-size:9pt;"><b>Giri� Ba�ar�l�</b></font><br /><br />
<font face="Verdana" style="font-size:9pt;"><i>Siteye Y�nlendiriliyorsunuz..</i> </font><br />
<img src="../imgs/18-1.gif" />
<meta http-equiv="refresh" content="1;url=default.asp">
<%Response.End
Else %>
<meta http-equiv="refresh" content="1;url=default.asp">
<font face="Verdana" style="font-size:9pt;"><i>Kullan�c� Ad� veya Parola Yanl��..</i></font>
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
Kullan�c� Ad� ve �ifrenizle giri� yap�n�z. E�er y�netici de�ilseniz l�tfen <a href="../default.asp">t�klay�n</a><br />
<br /><form action="Login.asp?islem=login" method="POST" class="style1">
<table>
<tr><td colspan="2" align="center">Y�netim Sayfas� Administrator Giri�i</td></tr>
<tr><td>Kullan�c� Ad�</td><td><input type="text" maxlength="12" name="strAccountID" value="<%=strAccountID%>"/></td></tr>
<tr><td>�ifre</td><td><input type="password"maxlength="12" name="strPasswd" /></td></tr>
<tr><td>G�velik Kodu</td><td><input type="text" name="SecurityCode" id="SecurityCode" maxlength="8"> </td>
<td><img id="simage" src="../securityImage.asp">
<a href="#" onClick="reloadImage();return false;" style="position:relative;top:-11px">
<label for="SecurityCode">( Kodu Yenile )</label></a></td></tr>
<tr><td colspan="2"><input type="submit" value="Giri�" style="width:300px"/></td></tr>
</table>
</form>
<%Else
Response.Redirect("Default.asp")
End If

End If %>