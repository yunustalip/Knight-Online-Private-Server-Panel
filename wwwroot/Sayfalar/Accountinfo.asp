<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<!--#include file="../guvenlik.asp"-->
<br><img src="imgs/AccountInfo.gif"><br><%ips=request.ServerVariables("REMOTE_HOST")
Response.Charset = "iso-8859-9"
response.expires=0
Dim MenuAyar
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='AccountInfo'")
If MenuAyar("PSt")=1 Then
if Session("login")="ok" Then
guncelle=trim(secur(Request.Querystring("guncelle")))
	If guncelle="bilgi" Then
	Set guncelle = Server.CreateObject("ADODB.Recordset")
	SQL = "Select * From TB_USER Where strAccountID='"&Session("username")&"'"
	guncelle.open SQL,Conne,1,3
	guncelle("stremail")=secur(request.Form("mail"))
	guncelle("gizlisoru")=secur(request.form("gsoru"))
	guncelle("cevap")=secur(request.Form("cevap"))
	guncelle.update
	guncelle.close
	set guncelle=nothing
	End If
set info=Conne.Execute("select * from tb_user where straccountid='"&Session("username")&"'")
password=len(info("strpasswd"))
email=info("strEmail")
cash=info("cashpoint")
gizlisoru=info("gizlisoru")
cevap=info("cevap")

%>
<script language="javascript">
function formpost(url,formid){
$.ajax({
   type: 'POST',
   url: url,
   data: $('#'+formid).serialize() ,
   start:  $('#ortabolum').html('<center>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif><br>G�ncelleniyor...</center>'),
   success: function(ajaxCevap) {
      $('#ortabolum').html(ajaxCevap);
   }
});
}
</script>
<center>
    <% if len(email)=0 or len(gizlisoru)=0 or len(cevap)=0 Then
	Response.Write "Bilgileriniz Eksik. L�tfen tamamlay�n�z<br><b>Bilgileri G�ncelle</b>"%>
    <form action="javascript:formpost('sayfalar/accountinfo.asp?guncelle=bilgi','addinfo')" method="post" id="addinfo"><table>
    <tr>
    <td ><strong>E-Mail</strong></td>
    <td>:</td>
    <td><input type="text" value="<%=email%>" name="mail" style="background-color:#FFFFFF; border-style: groove"></td>
    </tr>
    <tr>
      <td><strong>Gizli Soru</strong></td>
      <td>:</td>
      <td><select name="gsoru" style="background-color:#FFFFFF; border-style: groove">
	<%
	for x=1 to 8
	Response.Write "<option value="""&x&""">"&gizlis(""&x&"")&"</option>"
	next
	%></select></td>
    </tr>
    <tr>
      <td><strong>Cevap</strong></td>
      <td>:</td>
      <td><input type="text" value="<%=cevap%>" name="cevap" style="background-color:#FFFFFF; border-style: groove"></td>
    </tr>
    <tr>
      <td colspan="3" align="center"><input name="G�nder" type="submit" value="G�ncelle" style="background-image:url(imgs/layout_38.gif); color:#990000;border-collapse: inherit; border-style: groove;background:url(imgs/layout_38.gif)"></td>
      </tr>
    </table></form>
	<%
	else
%>
</center>
<br />
<br /><table   border="1" align="center" bordercolor="#CCCC00" style=" border-collapse: collapse;">
<tr>
<td colspan="3" align="center"><strong>Hesap Bilgileri - MyKOL</strong></td>
</tr>
<tr>
<td colspan="3">&nbsp;</td></tr>
  <tr>
    <td align="center"><strong><u>�ifre</u></strong></td>
    <td align="center"><strong><u>E-Mail</u></strong></td>
    <td align="center"><b>Gizli Soru &amp; Cevap</b></td>
    </tr>
  <tr>
    <td align="center" ><%for p=1 to password
	Response.Write "*"
	next%>&nbsp;<strong><a href="#" onclick="pageload('Sayfalar/accountinfo.asp?cat=accountinfo&change=pass','1');return false" class="link1"><br>(De�i�tir)</a></strong></td>
    <td align="center"><%=secur(email)%>&nbsp;<strong><br><a href="#" onclick="pageload('Sayfalar/accountinfo.asp?cat=accountinfo&change=email','1');return false" class="link1">(De�i�tir)</a></strong></td>
    <td align="center"><strong><a href="#" onclick="pageload('Sayfalar/accountinfo.asp?cat=accountinfo&change=secretquestion','1');return false" class="link1">(De�i�tir)</a></strong></td>
  </tr>
  <tr>
    <td colspan="3">&nbsp;</td>
    </tr>
  <tr>
    <td align="center"><strong>Cash Point</strong></td>
    <td align="center"><strong>Ticket Y�netimi</strong></td>
    <td align="center"></td>
  </tr>
  <tr>
    <td align="center"><%=cash%></td>
    <td align="center">&nbsp;&nbsp;&nbsp;<a href="#" onclick="pageload('Sayfalar/accountinfo.asp?cat=accountinfo&ticket=kontrol','1');return false" class="link1">Ticket Kontrol</a>&nbsp;/&nbsp;<a href="#" onclick="pageload('Sayfalar/submitticket.asp','1');return false" class="link1">Ticket G�nder</a>&nbsp;&nbsp;&nbsp;</td>
    <td align="center"></td>
  </tr>
</table>

	<% ticket=trim(secur(Request.Querystring("ticket")))
	if ticket="kontrol" Then
	set tickntrl=Conne.Execute("select * from tickets where charid='"&Session("username")&"'")
	%>
    <hr>
    <table width="397" border="0" align="center">
	<th colspan="3"><font color="#000000" size="2">Ticket Kontrol</font></th>
  <tr><%if not tickntrl.eof Then %>
    <td width="115"><strong>G�nderen</strong></td>
    <td width="5"><strong>:</strong></td>
    <td width="253"><%=tickntrl("charid")%></td>
  </tr>
  <tr>
    <td><strong>Konu</strong></td>
    <td><strong>:</strong></td>
    <td><%=tickntrl("subject")%></td>
  </tr>
  <tr>
    <td><strong>Mesaj</strong></td>
    <td><strong>:</strong></td>
    <td><%=tickntrl("message")%></td>
  </tr>
  <tr>
    <td><strong>G�nderim Zaman�</strong></td>
    <td><strong>:</strong></td>
    <td><%=tickntrl("date")%></td>
  </tr>
  <tr>
    <td valign="middle"><strong>Durum</strong></td>
    <td valign="middle"><strong>:</strong></td>
    <td valign="middle"><%if tickntrl("durum")="2" Then
	Response.Write "Okundu !<img src='imgs/accept.gif' align='absmiddle'>"
	%></td>
  <tr>
    <td><strong>Cevap</strong></td>
    <td><strong>:</strong></td>
    <td><%=tickntrl("cevap")%></td>
  </tr>
  <tr>
  <td></td>
  <td></td>
  <td align="center" valign="middle"> 
    <div align="center"><a href="#" onclick="pageload('Sayfalar/accountinfo.asp?cat=accountinfo&ticket=sil','1');return false" title="Ticketi Sil"><img src="imgs/Mail_delete.gif" alt="Ticketi Sil" height="32" border="0" align="absmiddle"><br>
      S�L</a></div>
  </td>
  </tr>
	<%
	elseif tickntrl("durum")="1" Then
	Response.Write "(Cevap)Bekleniyor...<img src='../imgs/Mail_reply.gif' align='absmiddle'></td></tr>"
	else
	End If
	else
	Response.Write "<tr><td>Ticket Bulunamad�.</td></tr>"
	End If
	%>
</table>
<hr>
	<%
	elseif ticket="sil" Then
	set ticketsill=Conne.Execute("delete tickets where charid='"&Session("username")&"'")
	End If
	change=trim(secur(Request.Querystring("change")))
	select case change
	case "secretquestion"
	%><br>
<form action="javascript:formpost('sayfalar/accountinfo.asp?change=sorudegistir','gizlisorudegisim')" method="post" id="gizlisorudegisim" name="gizlisorudegisim">
    <table>
    <tr>
    <td colspan="2" align="center" background="imgs/menubg.gif" class="style1">Gizli Soru De�i�imi</td>
    </tr>
    <tr>
    <td class="style3"><b>Gizli Soru:</b></td>
    <td><b><%=gizlis(gizlisoru)%></b></td>
    </tr>
    <tr>
      <td  class="style3"><b>Gizli Cevap:</b></td>
      <td><input type="text" name="gizlicevap"></td>
    </tr>
    <tr>
    <td  class="style3"><b>Yeni Gizli Soru :</b></td>
    <td><select name="yenisoru" class="styleform">
	<%
	for x=1 to 8
	Response.Write "<option value="""&x&""">"&gizlis(""&x&"")&"</option>"
	next
	%>
  </select></td></tr>
    <tr>
      <td  class="style3"><b>Yeni Gizli Cevap :</b></td>
      <td><input type="text" name="yenicevap"></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type="submit" value="De�i�tir" class="styleform"></td>
    </tr>
    </table>
    </form>
    <% case "sorudegistir"%><br>
	<br>
	<%gizlicevap=secur(request.form("gizlicevap"))
	yenisoru=secur(request.form("yenisoru"))
	yenicevap=secur(request.form("yenicevap"))
	if not gizlicevap="" or yenisoru="" or yenicevap="" Then
	if isnumeric(yenisoru)=false Then
	Response.End
	End If
	if yenisoru>0 and yerisoru<9 Then
	Set sorudegis = Server.CreateObject("ADODB.Recordset")
	SQL = "Select * From TB_USER Where strAccountID='"&Session("username")&"' and cevap='"&gizlicevap&"'"
	sorudegis.open SQL,Conne,1,3
	if not sorudegis.eof Then
	sorudegis("gizlisoru")=yenisoru
	sorudegis("cevap")=yenicevap
	sorudegis.update
	sorudegis.close
	set sorudegis=nothing
	Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&Session("username")&", nickli karakterin Gizli Sorusu "&yenisoru&", Gizli Cevab� "&yenicevap&" olarak de�i�tirildi.','"&now&"')")
	Response.Write "Gizli Soru De�i�tirildi !"
	else
	Response.Write("Yanl�� Cevap ! <a href='javascript:history.back(-1)'>Geri D�n</a>")
	End If
	else
	Response.Write("Yanl�� bilgi !!!")
	End If
	else
	Response.Write "Bo� alan b�rakmay�n�z <a href='javascript:history.back(-1)'>Geri D�n</a>"
	End If

	case "email"%><br>
	<form action="javascript:formpost('sayfalar/accountinfo.asp?change=emailok','emaildegisim')" method="post" id="emaildegisim" name="emaildegisim">
    <table align="center">
    <tr>
    <td colspan="2" align="center" background="imgs/menubg.gif" class="style1"><b>E-mail De�i�imi</b></td>
    </tr>
    <tr>
    <td width="100"><b>Gizli Soru : </b></td>
    <td ><strong><%=gizlis(gizlisoru)%></strong></td>
    </tr>
    <tr>
    <td><b>Gizli Cevap:</b></td>
    <td><input type="password" name="gizlicevap"></td>
    </tr>
    <tr>
    <td><b>Yeni E-Mail:</b></td>
    <td><input type="text" name="newemail"></td>
    </tr>
    <tr>
    <td colspan="2" align="center"><input type="submit" value="De�i�tir" class="styleform"></td>
    </tr>
    </table>
    </form>
    <% case "emailok"%>
	<br><br>
	<%
	gizlicevap=secur(request.form("gizlicevap"))
	newemail=secur(request.form("newemail"))
	
	if gizlicevap="" or newemail="" Then
	Response.Write "Bo� B�rakt���n�z alanlar var.<a href = ""javascript:history.back()""> Geri D�n </a>"
	else
	set changemail=Conne.Execute("select * from tb_user where straccountid='"&Session("username")&"' and cevap='"&gizlicevap&"'")
	if not changemail.eof Then
	set mailk=Conne.Execute("select * from tb_user where stremail='"&newemail&"'")
	if not mailk.eof Then
	Response.Write "<br><b>Girmi� oldu�unuz E-Mail Adresi Sistemimizde Kay�tl�d�r L�tfen Farkl� Bir Mail Adresi Giriniz</b>"
	Response.End
	End If
	set emailchange=Conne.Execute("update tb_user set stremail='"&newemail&"' where straccountid='"&Session("username")&"'")
	Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&Session("username")&", nickli karakterin E-Mail Adres, "&newemail&" olarak de�i�tirildi.','"&now&"')")
	Response.Write("E-Mail De�i�tirilmi�tir<br>Yeni E-Mail : "&secur(newemail))
	else
	Response.Write("Gizli Cevap Yanl�� L�tfen Tekrar Deneyin. <a href = ""javascript:history.back()""> Geri D�n </a>")
	End If

	End If

 case "pass" %>
 <br>
<form action="javascript:formpost('sayfalar/accountinfo.asp?change=passonay','sifredegisimi')" method="post" id="sifredegisimi" name="sifredegisimi">
<table width="300" border="0" align="center">
  <tr>
  <td colspan="2" align="center" background="imgs/menubg.gif" class="style1"><b>�ifre De�i�imi</b></td>
  </tr>
  <tr>
    <td width="104"><strong>Gizli Soru:</strong></td>
    <td width="146"><strong><%=gizlis(gizlisoru)%></strong></td>
  </tr>
    <tr>
    <td width="104"><strong>Cevap</strong></td>
    <td width="146"><input name="cevap" type="text" autocomplete="off" ></td>
  </tr>
  
  <tr>
    <td><strong>Yeni �ifre:</strong></td>
    <td><input name="newpass" type="password"/></td>
  </tr>
  <tr>
  <td ><strong>�ifre Tekrar:</strong></td>
  <td ><input name="repass" type="password" /></td>
  </tr>
  <tr>
  <td colspan="2" align="right"><input name="G�nder" type="submit" value="G�nder" class="styleform" onclick="this.value='��leminiz ger�ekle�tiriliyor.';this.form.submit()"/></td>
  </tr>
</table>
</form>
<% case "passonay"
cevap=trim(secur(request.form("cevap")))
newpass=trim(secur(request.form("newpass")))
repass=trim(secur(request.form("repass")))
if cevap="" or newpass="" or repass="" Then
Response.Write("<br><b>Bo� alan b�rakmay�n�z. <a href = ""#"" onclick=""pageload('Sayfalar/accountinfo.asp?cat=accountinfo&change=pass','1');return false""> Geri D�n </a></b>")
else
Set passcheck = Server.CreateObject("ADODB.Recordset")
SQL = "Select * From TB_USER Where strAccountID='"&Session("username")&"'"
passcheck.open SQL,Conne,1,3

if not cevap=passcheck("cevap") Then
Response.Write("<br><b>Yanl�� cevap l�tfen tekrar deneyin.<a href = ""#"" onclick=""pageload('Sayfalar/accountinfo.asp?cat=accountinfo&change=pass','1');return false""> Geri D�n </a></b>")
else
if not newpass=repass Then 
Response.Write("<br><b>�ifreler birbirini tutmuyor l�tfen tekrar deneyin.<a href = ""#"" onclick=""pageload('Sayfalar/accountinfo.asp?cat=accountinfo&change=pass','1');return false""> Geri D�n </a></b>")
else 
passcheck("strPasswd")=newpass
passcheck.update
tarih=cdate(now)
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&Session("username")&", nickli karakterin �ifresi "&newpass&" olarak de�i�tirildi.','"&now&"')")
Response.Write("<br>�ifreniz ba�ar�yla de�i�tirilmi�tir.")
End If
End If
End If

End select
End If

Else
Response.Write ("L�tfen kullan�c� giri�i yap�n�z.")
End If
Else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu b�l�m Admin taraf�ndan kapat�lm��t�r.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>