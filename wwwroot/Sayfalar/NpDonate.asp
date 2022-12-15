<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<!--#include file="../guvenlik.asp"-->
<%ips=Request.ServerVariables("REMOTE_HOST")
Response.Charset = "iso-8859-9"
Response.expires=0 
Response.Write "<base href=""http://"&Request.ServerVariables("server_name")&""">"
Dim MenuAyar,ksira,ips,islem,chars
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='NpDonate'")
If MenuAyar("PSt")=1 Then

if Session("login")="ok" Then%>
<br><img src="imgs/npdonate.gif"><br /><br /><br />
<%islem=secur(Request.Querystring("islem"))
if islem="" Then
islem=1
End If
if islem=1 Then
set chars=Conne.Execute("select * from account_char where straccountid='"&Session("username")&"'") 
if not chars.eof Then%>
<script language="javascript">
function npdonate(){
$.ajax({
   type: 'POST',
   url: 'sayfalar/npdonate.asp?islem=2',
   data: $('form#npbagis').serialize() ,
   start:  $('div#ortabolum').html('<br><br>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif><br>Sayfa YükLeniyor. Lütfen bekleyin.'),
   error: function(){
$('div#ortabolum').html('<br><br>&nbsp;&nbsp;&nbsp;<br>Hata oluþtu. Sayfa GörüntüLenemiyor...');
},
   success: function(ajaxCevap) {
      $('div#ortabolum').html(ajaxCevap);
   }
});
}

</script>
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

<form action="sayfalar/npdonate.asp?islem=2" onsubmit="javascript:npdonate(); return false;" method="post" id="npbagis" name="npbagis">
<table>
<tr><td class="txt">Karakter Adý :</td>
<td><%Response.Write "<select name=""charname"" class=""inpt"" style=""padding-top:2px"">"
If Len(trim(chars("strcharid1")))>0 Then
Response.Write "<option value=""1"" style=""height:15px"">"&chars("strcharid1")&"</option>"
End If
If Len(trim(chars("strcharid2")))>0 Then
Response.Write "<option value=""2"" style=""height:15px"">"&chars("strcharid2")&"</option>"
End If
If Len(trim(chars("strcharid3")))>0 Then
Response.Write "<option value=""3"" style=""height:15px"">"&chars("strcharid3")&"</option>"
End If
Response.Write "</select>"%>
</td></tr>
<tr>
<td class="txt"><strong>Miktar :</strong></td>
<td><input type="text" name="npmiktar" id="npmiktar" class="inpt"></td>
</tr>
<tr>
<td colspan="2" align="right"><input type="submit" value="Np Baðýþla" class="styleform" style="color:#8E6400;font-weight:bold;font-size:10px;" onclick="if (this.form.npmiktar.value!=''){return confirm('Clanýnýza '+document.getElementById('npmiktar').value+' Np baðýþlamak istiyormusunuz?')}else{alert('Boþ Býrakmayýnýz!');return false}"></td>
</tr></table>
</form>
<%
End If
End If
If islem=2 Then
Dim charname,npmiktar
charname=secur(Trim(Request.Form("charname")))
npmiktar=request.form("npmiktar")
If IsNumeric(npmiktar)=False Or IsNumeric(charname)=False Then
Response.Write "<div class=""errortxt"">Sayýsal Deðerler Giriniz.</div>"
Response.End
End If

If npmiktar<=0 Then
Response.Write "<div class=""errortxt"">Clana En Az 1 Np Baðýþlanabilir.</div>"
Response.End 
End If

If npmiktar>10000000 Then
Response.Write "<div class=""errortxt"">Lütfen daha küçük bir deðer giriniz.</div>"
Response.End 
Else
npmiktar=clng(int(npmiktar))
End If

If charname="1" or charname="2" or charname="3" Then
Else
Response.Redirect("NpDonate.Asp")
Response.End
End If

Dim charkontrol
Set charkontrol=Conne.Execute("Select strCharID"&charname&" from ACCOUNT_CHAR Where strAccountID='"&Session("username")&"'")

If Not charkontrol.Eof Then
charname=charkontrol(0)
Dim user,sql
Set user = Server.CreateObject("ADODB.Recordset")
sql ="Select Struserid, Knights, Loyalty From Userdata Where StrUserID='"&charname&"'"
user.Open sql,conne,1,3

If npmiktar>user("loyalty") Then
Response.Write "<div class=""errortxt"">HATA: Baðýþlanan Miktar Mevcut Npnizden Daha Fazladýr.<br><br>Karakterinizdeki Mevcut National Point: "&user("loyalty")&"</div>"
Response.End
Else

If user("knights")=0 Then
Response.Write "<div class=""errortxt"">HATA: Clanýnýz bulunmamaktadýr.</div>"
Response.End
End If

Dim clankntrl,cha,onlinek
Set clankntrl=Conne.Execute("select idnum from knights where idnum="&user("knights")&"")

If clankntrl.eof Then
Response.Write "<div class=""errortxt"">Clan Bulunamadý!</div>"
Response.End
End If

Set cha = Server.CreateObject("ADODB.Recordset")
sql ="select * from npdonate where userid='"&charname&"' and clan="&user("knights")&""
cha.open sql,conne,1,3

Set onlinek=Conne.Execute("select strcharid from currentuser where strcharid='"&charname&"'")

If onlinek.Eof Then

If cha.Eof Then
Conne.Execute("insert into npdonate (userid,np,clan) values('"&charname&"','"&npmiktar&"','"&user("knights")&"')")
user("loyalty")=user("loyalty")-npmiktar
user.update
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&charname&" nickli karakteri "&user("knights")&" numaralý clana "&npmiktar&" Np baðýþladý.Kalan Np:"&user("loyalty")&"','"&now&"')")
Response.Write "<br><b><font class=style4>Iþlem Baþarýlý !</font></b><br><br>"
Response.Write "Clanýnýza "&npmiktar&" Np Baðýþlandý.<br>"
Response.Write "Kalan Np: "&user("loyalty")

Else

cha("np")=clng(cha("np"))+npmiktar
cha.Update
cha.Close
Set cha=Nothing
user("loyalty")=user("loyalty")-npmiktar
user.Update

Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&charname&" nickli karakteri "&user("knights")&" numaralý clana "&npmiktar&" Np baðýþladý.Kalan Np:"&user("loyalty")&"','"&now&"')")
Response.Write "<br><br><b><font class=""style4"">Iþlem Baþarýlý !</font></b><br><br>"
Response.Write "Clanýnýza "&npmiktar&" Np Baðýþlandý.<br>"
Response.Write "Kalan Np: "&user("loyalty")
End If

Else
Dim logadd,sqlx
Set logadd = Server.CreateObject("ADODB.Recordset")
sqlx ="select * from nplog"
logadd.Open sqlx,conne,1,3

logadd.Addnew
logadd("struserid")=charname
logadd("clanno")=user("knights")
logadd("np")=npmiktar
logadd("durum")="verdi"
logadd.Update
logadd.Close
Set logadd=Nothing

Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&charname&" Nickli karakteri "&user("knights")&" numaralý clana "&npmiktar&" Np baðýþladý.','"&now&"')")
Response.Write "<div class=""errortxt""><br><b>Np Baðýþlandý oyundan çýkýp tekrar girdiðinizde güncellenecek.</b></div>"
End If

End If
Else
Response.Write "<div class=""errortxt"">Karakter Bulunamadý!</div>"
End If
End If

Else
Response.Write "<div class=""errortxt"">Lütfen kullanýcý giriþi yapýnýz.</div>"
End If
else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>