<!--#include file="_inc/conn.asp"-->
<!--#include file="function.asp"-->
<%Response.expires=0
Dim MenuAyar,ksira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='PowerUpStore'")
If MenuAyar("PSt")=1 Then
cashban=Session("cashban")
if cashban="" Then
Session("cashban")=0
elseif cashban>=5 Then
Conne.Execute("update tb_user set strauthority=255 where straccountid='"&Session("username")&"'")
Session("username")=""
Session("login")=""
Session("yetki")=""
Session.abandon
Response.write "<script>alert('Hesab�n�z Bloke Olmu�tur. Admin �le �leti�ime Ge�iniz');self.close()</script>"
Response.End
End If%>

<style type="text/css">
<!--
.button {
	background-color: #C00;
	font-family: "MS Serif", "New York", serif;
	font-size: 14px;
	color: #FFF;
	font-weight: bold;
}
-->
</style><%if Request.Querystring("control")="recharge2" Then 
cashcode1=request.form("cashcode1")
cashcode2=request.form("cashcode2")
cashcode3=request.form("cashcode3")
cashcode4=request.form("cashcode4")
cashcode5=request.form("cashcode5")

if isnumeric(cashcode1)=false or isnumeric(cashcode2)=false or isnumeric(cashcode3)=false or isnumeric(cashcode4)=false or isnumeric(cashcode5)=false Then
Response.Write "L�tfen Do�ru Cash Code Giriniz"
Response.End 
End If

if cashcode1="" or cashcode2="" or cashcode3="" or cashcode4="" or cashcode5="" Then
Response.Redirect("powerupstore.asp?control=recharge")
Response.End
End If
cashcode= cashcode1&"-"&cashcode2&"-"&cashcode3&"-"&cashcode4&"-"&cashcode5
set cashkontrol=Conne.Execute("select * from cash_table where cashcode='"&cashcode&"' and durum='on' ")
if cashkontrol.eof Then
Session("cashban")=Session("cashban")+1
if Session("cashban")>=5 Then
Conne.Execute("update tb_user set strauthority=255 where straccountid='"&Session("username")&"'")
Session("username")=""
Session("login")=""
Session("yetki")=""
Session.abandon
Response.write "<script>alert('Hesab�n�z Bloke Olmu�tur. Admin �le �leti�ime Ge�iniz');self.close()</script>"
Response.End
End If
Response.Write("<center>Yanl�� Cash Code girdiniz.<br>")
Response.Write 5-Session("cashban")&" Kez daha �st �ste yanl�� �ifre girerseniz hesab�n�z bloke olacakt�r!"
elseif not cashkontrol.eof Then
Conne.Execute("update cash_table set durum='off', alanchar='"&Session("username")&"' where cashcode='"&cashcode&"' ")
Conne.Execute("update tb_user set cashpoint=cashpoint+"&cashkontrol("cashmiktar")&" where strAccountID='"&Session("username")&"'")
Response.Write "<center>Hesab�n�za "&cashkontrol("cashmiktar")&" Cash Point Eklenmi�tir</center>"
End If
%>


<%else%>
<form id="form1" name="form" method="post" action="powerupstore.asp?control=recharge2">
<table width="452" border="0" align="center">
  <tr>
    <td height="33" colspan="2" align="center"><p><strong>Cash Code Registration</strong></p></td>
  </tr>
  <tr>
    <td height="50"><label for="cashcode">Cash Code :</label>
    </td>
    <td>
      <input name="cashcode1" type="text" id="cashcode1" size="5" maxlength="4" />-<input name="cashcode2" type="text" id="cashcode2" size="5" maxlength="4" />-<input name="cashcode3" type="text" id="cashcode3" size="5" maxlength="4" />-<input name="cashcode4" type="text" id="cashcode4" size="5" maxlength="4" />-<input name="cashcode5" type="text" id="cashcode5" size="5" maxlength="4" />
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input name="button" type="submit" class="button" id="button" value="Code Register" /></td>
  </tr>
</table></form>
<%End If
else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu b�l�m Admin taraf�ndan kapat�lm��t�r.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>
