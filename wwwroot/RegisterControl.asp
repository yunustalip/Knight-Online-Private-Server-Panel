<!--#include file="_inc/conn.asp"-->
<!--#include file="function.asp"-->
<%dim username,pwd,pwd2,email,sCAPTCHA,gizlisoru,gizlicevap,menuayar
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Register'")
Response.Charset = "iso-8859-9"
If MenuAyar("PSt")=1 Then
ip=Request.ServerVariables("REMOTE_ADDR")


 username = tr(trim(secur(request.form("username"))))
 pwd = trim(secur(request.form("pwd")))
 pwd2 = trim(secur(request.form("pwd2")))
 email = trim(emailAddressValidation(request.form("email")))
 sCAPTCHA = trim(ucase(secur(Request.Form("formHC"))))
 gizlisoru= trim(secur(request.form("gizlisoru")))
 gizlicevap=trim(secur(request.form("gizlicevap")))

 if username="" or pwd="" or pwd2="" or email="" or gizlisoru="" or gizlicevap="" Then
 Response.Write("<script>alert('Boþ Býraktýðýnýz Alanlar Var. Lütfen doldurunuz !')</script>")
 elseif len(username)>20 Then
 Response.Write "<script>alert('Kullanýcý adý 20 karakterden fazla olamaz.')</script>"
 elseif len(pwd)>13 or len(pwd2)>13 Then
 Response.Write "<script>alert('Parolanýz 13 karakterden büyük olamaz.')</script>"
 elseif len(email)>50 Then
 Response.Write "<script>alert('E-Mail Adresiniz 50 karakterden büyük olamaz.')</script>"
 elseif len(gizlisoru)>100 or len(gizlisoru)>100 Then
 Response.Write "<script>alert('Gizli Soru veya cevabýnýz max. 100 karakter olabilir.')</script>"
 elseif EmailKontrol(email)="False" Then
 Response.Write "<script>alert('Lütfen geçerli bir e-mail adresi girin.')</script>"
  elseif pwd<>pwd2 Then 
 Response.Write("<script>alert('Þifreler Uyuþmuyor. Tekrar Deneyin')</script>")
 elseif pwd=pwd2 Then 
 if NOT sCAPTCHA = Session("human-control_" & Session.SessionID) or sCAPTCHA = "" Then
 Response.Write "<font face=Verdana style=font-size:10pt; color=red><b>Yanlýþ Kod !</b></font>" : Response.End
 else

Set userreg = Server.CreateObject("ADODB.Recordset")
sql = "Select * From tb_user where strAccountID='"&username&"'"
userreg.open sql,conne,1,3

if not userreg.eof Then 
Response.Write ("<script>alert('Bu Kullanýcý Kayýtlýdýr. Lütfen Baþka Bir Kullanýcý Adý Yazýn.')</script>")
else

set emailkontr=Conne.Execute("Select * From tb_user where stremail='"&email&"'")
if not emailkontr.eof Then
Response.Write ("<script>alert('Bu E-Mail Kayýtlýdýr. Lütfen Baþka Bir E-Mail Yazýn.')</script>")
else

gizlisoru=gizlis(gizlisoru)

userreg.Addnew
userreg("strAccountID")=username
userreg("strPasswd")=pwd
userreg("idays")="6"
userreg("strSocNo")="1"
userreg("stremail")=email
userreg("gizlisoru")=gizlisoru
userreg("cevap")=gizlicevap
userreg.update
userreg.close
set userreg=nothing
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ip&"','"&username&" : "&pwd&" Kayýt oldu','"&now&"')")

%>
<script>
document.getElementById('kayit').reset;
$('#ortabolum').html('<img src="imgs/register.gif" /><br /><br />
<center><br />
  <hr />
  <b>Kullanýcý Bilgileri</b><br />
  <br />
<b>Kullanýcý Adý : </b> <%=username%><br>
<b>Þifre : </b> <%for x=1 to len(pwd)
Response.Write "*"
next 
%><br>
<b>E-mail : </b><%=email%>
<br>
<br>
Kayýt Olduðunuz için teþekkürler !
<hr>
</center>');
</script>
<% Session("human-control_" & Session.SessionID)=""
End If
End If 
End If 
End If 

else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
end  if
MenuAyar.Close
Set MenuAyar=Nothing%>