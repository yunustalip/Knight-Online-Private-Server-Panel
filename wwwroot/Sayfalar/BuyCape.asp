<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<!--#include file="../guvenlik.asp"-->
<%response.expires=0
Dim MenuAyar
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='BuyCape'")
If MenuAyar("PSt")=1 Then
if Session("login")="ok" Then
ips=Request.ServerVariables("REMOTE_HOST")%>
<script language="javascript">
function pelerinal(){
$.ajax({
   type: 'POST',
   url: 'sayfalar/buycape.asp?islem=2',
   data: $('form#form1').serialize() ,
   start:  $('div#ortabolum').html('<br><br><br><br>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif>'),
   error: function(){
$('div#ortabolum').html('&nbsp;&nbsp;&nbsp;<br>Hata oluþtu. Sayfa Görüntülenemiyor...');
},
   success: function(ajaxCevap) {
      $('div#ortabolum').html(ajaxCevap);
   }
});
}
</script>
<br>
<img src="../imgs/buycape.gif" alt=""/><br><br><br><%islem=secur(Request.Querystring("islem"))
if islem="" Then
islem="1"
End If
select case islem
case "1"
set users=Conne.Execute("select * from ACCOUNT_CHAR where straccountid='"&Session("username")&"'")
if not users.eof Then%>
<span style="color:#8E6400;font-weight:bold">Pelerin alabilmek için kasanýzda en az  1.000.000.000 coins bulunmalýdýr. <br>

Iþlemlerinizi gerçekleþtirmeden önce clanýnýzdaki bütün userlerin offline olmasýna dikkat edin.
</span><br><br>
<style>
.inpt{
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:10px;
border:solid 1px;
border-color:#8E6400;
color:#8E6400;
font-weight:bold;
height:20px;
text-decoration:inherit;
background-color:#F4F4F4
}
.txt{
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:10px;
color:#8E6400;
font-weight:bold;
}
</style>
<form action="javascript:pelerinal();" method="post" id="form1" name="form1">
<span style="color:#8E6400;font-weight:bold">Karakterinizi seçiniz: </span>
<select name="kulad" class="inpt" style="padding-top:2px">
<% if len(trim(users("strcharid1")))>0 Then %>
<option value="<%=users("strcharid1")%>" style="height:15px"><%=users("strcharid1")%></option>
<% End If
 if len(trim(users("strcharid2")))>0 Then  %>
 <option value="<%=users("strcharid2")%>" style="height:15px"><%=users("strcharid2")%></option>
 <% End If
 if len(trim(users("strcharid3")))>0 Then%>
 <option value="<%=users("strcharid3")%>" style="height:15px"><%=users("strcharid3")%></option>
 <% End If %>
</select><br><br />
 <input type="submit" value="Pelerin Al >>" class="styleform" style="color:#8E6400;font-weight:bold;font-size:10px;">
</form>

<%End If

case "2"
set users=Conne.Execute("select * from ACCOUNT_CHAR where straccountid='"&Session("username")&"'")
if not users.eof Then
kulad=secur(request.form("kulad"))
if kulad=users("strcharid1") or kulad=users("strcharid2") or kulad=users("strcharid3") Then
set clan=Conne.Execute("select * from userdata where struserid='"&kulad&"'")
if not clan.eof Then
if not clan("knights")="0" and clan("fame")="1" Then
set clan2=Conne.Execute("select * from knights where idnum='"&clan("knights")&"'")
if not clan2.eof Then
if clan2("chief")=kulad Then
if clan2("points")>=360000 Then

pelerinfiyat=1000000000

if clan("gold")>=pelerinfiyat Then
set paraal=Conne.Execute("update userdata set gold=gold-'"&pelerinfiyat&"' where struserid='"&kulad&"'")
Set pelerinal= Server.CreateObject("ADODB.Recordset")
cSQL = "Select * From KNIGHTS Where IDNum='"&clan2("idnum")&"'"
pelerinal.open cSQL,Conne,1,3
pelerinal("scape")=0
pelerinal.update
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&Session("username")&" Hesabýnýn "&trim(kulad)&" karakteri tarafýndan "&trim(clan2("idname"))&" Clanýna pelerin Alýndý.','"&now&"')")
Response.Write "<font size=""2""><br><b>"&clan2("idname")&"</b> clanýna baþarýyla pelerin alýnmýþtýr !</font>"
users.close
set users=nothing
clan.close
set clan=nothing
clan2.close
set clan2=nothing
pelerinal.close
set pelerinal=nothing
else
Response.Write "<br><br><b>Pelerin alabilmeniz için kasanýzda en az "&pelerinfiyat&" bulunmasý gerekmektedir.</b>"
End If
else
Response.Write "<br><br><b>Pelerin alabilmeniz için clanýnýzýn en az Grade 3 olmasý gerekmektedir.</b>"
End If
else
Response.Write "<br><br><b>Clanýn lideri siz deðilsiniz.</b>"
End If
else
Response.Write "<br><br><b>Clan bulunamadý</b>"
End If
else
Response.Write "<br><br><b>Clan lideri deðilsiniz !</b>"
End If

End If

End If

End If
end select
End If 


else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
end  if
MenuAyar.Close
Set MenuAyar=Nothing%>