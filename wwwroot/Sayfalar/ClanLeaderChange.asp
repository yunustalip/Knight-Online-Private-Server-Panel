<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<!--#include file="../guvenlik.asp"-->
<%response.expires=0
if Session("login")="ok" Then
Dim MenuAyar
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='ClanLeaderChange'")
If MenuAyar("PSt")=1 Then
ips=Request.ServerVariables("REMOTE_HOST")
%>
<script language="javascript">
function leaderchange1(){
$.ajax({
   type: 'POST',
   url: 'sayfalar/clanleaderchange.asp?islem=2',
   data: $('form#clanleaderchange').serialize() ,
   start:  $('div#ortabolum').html('<br><br><br><br>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif>'),
   error: function(){
$('div#ortabolum').html('&nbsp;&nbsp;&nbsp;<br>Hata oluþtu. Sayfa Görüntülenemiyor...');
},
   success: function(ajaxCevap) {
      $('div#ortabolum').html(ajaxCevap);
   }
});
}

function leaderchange2(){
$.ajax({
   type: 'POST',
   url: 'sayfalar/clanleaderchange.asp?islem=3',
   data: $('form#clanleaderchange2').serialize() ,
   start:  $('div#ortabolum').html('<br><br><br><br>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif>'),
   error: function(){
$('div#ortabolum').html('&nbsp;&nbsp;&nbsp;<br>Hata oluþtu. Sayfa Görüntülenemiyor...');
},
   success: function(ajaxCevap) {
      $('div#ortabolum').html(ajaxCevap);
   }
});
}

</script><style>
.txt{
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:10px;
color:#8E6400;
font-weight:bold;
}
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
</style><br><img src="imgs/clanleaderchange.gif" /><br /><br /><br />	
<%islem=secur(Request.Querystring("islem"))
if islem="" Then
islem="1"
End If

if islem="1" Then
set chars=Conne.Execute("select strcharid1,strcharid2,strcharid3 from account_char where straccountid='"&Session("username")&"' ")
set chrs=Conne.Execute("select struserid,fame,knights from userdata where struserid='"&chars("strcharid1")&"' and fame=1 and knights<>0 or struserid='"&chars("strcharid2")&"' and fame=1 and knights<>0 or struserid='"&chars("strcharid3")&"' and fame=1 and knights<>0 ")
%><form action="javascript:leaderchange1();" method="post" id="clanleaderchange" name="clanleaderchange">
<table >
<tr>
<td class="txt"><%if not chrs.eof Then%>Karakterinizi Seçiniz:</td>
<td><select name="charid" class="inpt" style="padding-top:2px">
<%do while not chrs.eof
if len(trim(chrs("struserid")))>0 Then
Response.Write "<option value='"&trim(chrs("struserid"))&"' style=""height:15px"">"&trim(chrs("struserid"))&"</option>"&vbcrlf
End If
chrs.movenext
loop
 %></select></td>
</tr>
<tr>
  <td colspan="2" align="center"><input type="submit" class="styleform" style="color:#8E6400;font-weight:bold;font-size:10px;" value="Devam Et >>"></td>
  </tr>
</table>
</form>
<%else
Response.Write "<div class=""errortxt"">Clan Lideri Karakteriniz Bulunmamaktadýr.</div></td></tr></table>"
End If
elseif islem="2" Then
charid=secur(request.form("charid"))
if charid="" Then
Response.Redirect("clanleaderchange.asp?islem=1")
End If

set charkntrl=Conne.Execute("select * from account_char where straccountid='"&Session("username")&"' and strcharid1='"&charid&"' or strcharid2='"&charid&"' or strcharid3='"&charid&"'")
if not charkntrl.eof Then

set clank=Conne.Execute("select * from userdata where struserid='"&charid&"'")
if clank("knights")="0" or clank("fame")<>"1" Then
Response.Write "<div class=""errortxt"">Clan Lideri Deðilsiniz !</div>"
Response.End
End If

set clank2=Conne.Execute("select * from knights where idnum='"&clank("knights")&"'")
if clank2.eof Then
Response.Write("<div class=""errortxt"">Clan Bulunamadý!</div>")
Response.End
End If
if not charid=trim(clank2("chief")) Then
Response.Write "<div class=""errortxt"">Clan Lideri Deðilsiniz !</div>"
Response.End
End If
set clanuye=Conne.Execute("select * from userdata where knights='"&clank("knights")&"'")

%>
<form action="javascript:leaderchange2();" method="post" id="clanleaderchange2" name="clanleaderchange2">
<table>
<tr>
<td class="txt">Clan Adý:</td>
<td><%=clank2("idname")%></td>
</tr>
<tr>
<td class="txt">Yeni Leaderi seçiniz:</td>
<td><select name="newleader" class="inpt" style="padding-top:2px">

<% if not clanuye.eof Then
do while not clanuye.eof
Response.Write "<option value="""&clanuye("struserid")&"""  style=""height:15px"">"&clanuye("struserid")&"</option>" 
clanuye.movenext
loop
End If
%>
</select><input type="hidden" value="<%=charid%>" name="oldleader"></td>
</tr>
<tr>
  <td colspan="2" align="center"><input type="submit" class="styleform" style="color:#8E6400;font-weight:bold;font-size:10px;" value="Claný Devret >>"></td>
  </tr>
</table>
</form>
<%

else
Response.Redirect("default.asp?cat=clanleaderchange&islem=1")
End If

elseif islem="3" Then
oldleader=secur(request.form("oldleader"))
newleader=secur(request.form("newleader"))

set charon=Conne.Execute("select * from currentuser where strcharid='"&oldleader&"' or strcharid='"&newleader&"'")

if not charon.eof Then
Response.Write "<div class=""errortxt"">Karakterler Oyunda Olmamalýdýr</div>"
Response.End
else
set charkntrl=Conne.Execute("select * from account_char where straccountid='"&Session("username")&"' and strcharid1='"&oldleader&"' or strcharid2='"&oldleader&"' or strcharid3='"&oldleader&"'")
if not charkntrl.eof Then

set clank=Conne.Execute("select * from userdata where struserid='"&oldleader&"' and authority<>255")
if not clank.eof Then
set newleaderc=Conne.Execute("select * from userdata where struserid='"&newleader&"' and authority<>255")
if not newleaderc.eof Then
set clank2=Conne.Execute("select * from knights where idnum='"&clank("knights")&"'")

if clank("knights")="0" or clank("fame")<>"1" Then
Response.Write("<div class=""errortxt"">Clan Lideri Deðilsiniz.</div>")
Response.End
End If

if not oldleader=trim(clank2("chief")) Then
Response.Write "<div class=""errortxt"">Clan Lideri Deðilsiniz !</div>"
Response.End
End If


if newleaderc("knights")=clank("knights") Then

Conne.Execute("update userdata set fame=5 where struserid='"&oldleader&"'")
Conne.Execute("update knights set ViceChief_1=NULL where ViceChief_1='"&newleader&"' and idnum='"&clank("knights")&"'")
Conne.Execute("update knights set ViceChief_2=NULL where ViceChief_2='"&newleader&"' and idnum='"&clank("knights")&"'")
Conne.Execute("update knights set ViceChief_3=NULL where ViceChief_3='"&newleader&"' and idnum='"&clank("knights")&"'")
Conne.Execute("update knights set chief='"&newleader&"' where idnum='"&clank("knights")&"'")
Conne.Execute("update userdata set fame=1, knights='"&clank("knights")&"' where struserid='"&newleader&"'")
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&trim(clank2("idname"))&" Clanýnýn Liderliði "&trim(oldleader)&" karakterinden "&newleader&" karakterine devrolmuþtur.','"&now&"')")
Response.Write "<br>"&(clank2("idname")&"Clanýnýn yeni lideri&nbsp;"&newleader&"&nbsp;olmuþtur.")

else
Response.Write "<div class=""errortxt"">Yeni leader clanýnýzda bulunmalýdýr.</div>"
End If
else
Response.Write "<div class=""errortxt"">Karakter Bulunamadý!</div>"
End If
else
Response.Write "<div class=""errortxt"">Karakter Bulunamadý!</div>"
End If
End If
End If
End If

	else
Response.Write "<div class=""errortxt"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</div>"
End If
MenuAyar.Close
Set MenuAyar=Nothing
		else 
	Response.Write ("Lütfen kullanýcý giriþi yapýnýz.")
	End If
%>