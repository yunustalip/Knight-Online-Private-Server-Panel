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
 Response.Write("<script>alert('Bo� B�rakt���n�z Alanlar Var. L�tfen doldurunuz !')</script>")
 elseif len(username)>20 Then
 Response.Write "<script>alert('Kullan�c� ad� 20 karakterden fazla olamaz.')</script>"
 elseif len(pwd)>13 or len(pwd2)>13 Then
 Response.Write "<script>alert('Parolan�z 13 karakterden b�y�k olamaz.')</script>"
 elseif len(email)>50 Then
 Response.Write "<script>alert('E-Mail Adresiniz 50 karakterden b�y�k olamaz.')</script>"
 elseif len(gizlisoru)>100 or len(gizlisoru)>100 Then
 Response.Write "<script>alert('Gizli Soru veya cevab�n�z max. 100 karakter olabilir.')</script>"
 elseif EmailKontrol(email)="False" Then
 Response.Write "<script>alert('L�tfen ge�erli bir e-mail adresi girin.')</script>"
  elseif pwd<>pwd2 Then 
 Response.Write("<script>alert('�ifreler Uyu�muyor. Tekrar Deneyin')</script>")
 elseif pwd=pwd2 Then 
 if NOT sCAPTCHA = Session("human-control_" & Session.SessionID) or sCAPTCHA = "" Then
 Response.Write "<font face=Verdana style=font-size:10pt; color=red><b>Yanl�� Kod !</b></font>" : Response.End
 else

Set userreg = Server.CreateObject("ADODB.Recordset")
sql = "Select * From tb_user where strAccountID='"&username&"'"
userreg.open sql,conne,1,3

if not userreg.eof Then 
Response.Write ("<script>alert('Bu Kullan�c� Kay�tl�d�r. L�tfen Ba�ka Bir Kullan�c� Ad� Yaz�n.')</script>")
else

set emailkontr=Conne.Execute("Select * From tb_user where stremail='"&email&"'")
if not emailkontr.eof Then
Response.Write ("<script>alert('Bu E-Mail Kay�tl�d�r. L�tfen Ba�ka Bir E-Mail Yaz�n.')</script>")
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
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ip&"','"&username&" : "&pwd&" Kay�t oldu','"&now&"')")

%>
<script>
document.getElementById('kayit').reset;
$('#ortabolum').html('<img src="imgs/register.gif" /><br /><br />
<center><br />
  <hr />
  <b>Kullan�c� Bilgileri</b><br />
  <br />
<b>Kullan�c� Ad� : </b> <%=username%><br>
<b>�ifre : </b> <%for x=1 to len(pwd)
Response.Write "*"
next 
%><br>
<b>E-mail : </b><%=email%>
<br>
<br>
Kay�t Oldu�unuz i�in te�ekk�rler !
<hr>
</center>');
</script>
<% Session("human-control_" & Session.SessionID)=""
End If
End If 
End If 
End If 

else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu b�l�m Admin taraf�ndan kapat�lm��t�r.</span></b>"
end  if
MenuAyar.Close
Set MenuAyar=Nothing%>