 <!--#include file="_inc/conn.asp"-->
<%Response.expires=0
Response.Charset = "iso-8859-9"
Dim MenuAyar,ksira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Login'")
If MenuAyar("PSt")=1 Then
%>
<table width="200" cellspacing="0" cellpadding="0" border="0" >
	<tr>
	<td background="imgs/sub_menu_title_bg.gif"  width="185" height="68" align="center" class="style1" style="padding-top:15px">
	<% if Session("login")="ok" Then Response.Write("Kullanýcý Menü&nbsp;") else Response.Write("Kullanýcý Giriþi&nbsp;")%>
	</td>
	</tr>
	<tr>
         <td background="imgs/sub_menu_bg.gif" style="padding-left: 10px;padding-top:-10px">
	<% if Session("login")="ok" Then 
	Dim uch
	Set uch =Conne.Execute("Select * From tb_user where strAccountID='"&Session("username")&"'")
	
	if not uch.eof Then

	if uch("strauthority")="255" Then
	with response
	.write "<font face=""arial,helvetica"" size=""2"">"
	.write "<p align=""center""><b>Giriþiniz Engellenmiþtir.</b><br><br>"
	.write "<a href=""javascript:loging()""><b>Geri dön</b></a></p>"
	.write "</font>"
	end with
	Session.abandon
	Response.End
	End If
if Session("yetki")="1" Then
%><script>
$(document).ready(function(){
  $("#komut").focus();
});
function komutlar(){
$('#kmt').slideToggle("fast")
}
function komutgir(kmt){
$('input#komut').val(kmt)
$("#komut").focus();
}
function komutyolla(){
$.ajax({
   type: 'get',
   url: 'Gmpage/Gamem.asp?user=gmkomut',
   data: 'komut='+$('#komut').val() ,
   success: function(ajaxCevap) {
      $('div#ortabolum').html(ajaxCevap);
   }
});
$('#komut').val('')
}

</script>
<form action="javascript:komutyolla();" method="get">
<div id="gmmen"><div id="gondr"></div>
Gm Menü: <input type="text" name="komut" id="komut" style="width:25%;background-color:#000;font-weight:bold;color:#00FF00;border-style:inset" autocomplete="off">
<input type="submit" value="Yolla" style="background-color:#000;color:#00FF00;border-style:groove">
<a onclick="komutgir('/kill ')" class="ylink">User Dc</a>|
<a onclick="komutgir('/open ')" class="ylink">Savaþ Aç(/Open)</a> |
<a onclick="komutgir('/open2 ')" class="ylink">Savaþ Aç(/Open2)</a> |
<a onclick="komutgir('/open3 ')" class="ylink">Savaþ Aç(/Open3)</a> |
<a onclick="komutgir('/close ')" class="ylink">Savaþý Kapat</a> |
<a onclick="komutgir('/permanent ');komutyolla();komutgir(prompt('Oyunda Kalan Premium Gününün yazýlý olduðu kýsmý deðiþtirmek istediðiniz yazýyý yazýn.'));komutyolla()" class="ylink">Permanent Gir</a>
</div>
</form>
<%End If%>
	<center class="style3">
	  Hoþgeldiniz <% =Session("username")%> 
	  </center>
	<br>
        <b><font color="#330099" style="margin-left:40px"><u>Karakterleriniz</u></font></b> &nbsp;<br>
	<%Dim accch,sql
	Set accch = Conne.Execute("Select * From ACCOUNT_CHAR where strAccountID='"&Session("username")&"'")
	Else
	Session("login")=""
	Response.Redirect("default.asp")
	Response.End
	End If
	If Not accch.eof Then
	Dim charid1,charid2,charid3,onlinechar
	charid1=trim(accch("strcharid1"))
	charid2=trim(accch("strcharid2"))
	charid3=trim(accch("strcharid3"))
	Session("charid1")=charid1
	Session("charid2")=charid2
	Session("charid3")=charid3
	Set OnlineChar = Conne.Execute("Select strcharid From currentuser where strAccountID='"&trim(Session("username"))&"' ")
	
	if len(charid1)>0 Then
	Response.Write("<center><a href=""Karakter-Detay/"&CharID1&""" onclick=""pageload('Karakter-Detay/"&CharID1&"');chngtitle('"&CharID1&" > Karakter Detay');return false"" class=""link2"" style=""display:block"">"&CharID1&"&nbsp;&nbsp;&nbsp;")
	If Not onlinechar.eof Then
	If Trim(onlinechar("strcharid"))=CharID1 Then
	Response.Write "<img src=""imgs/on.gif"" align=""absmiddle"" border=""0""></a></center>"
	else 
	Response.Write "<img src=""imgs/off.gif"" align=""absmiddle"" border=""0""></a></center>"
	End If
	Else
	Response.Write "<img src=""imgs/off.gif"" align=""absmiddle"" border=""0""></a></center>"
	End If
	End If

	If Len(charid2)>0 Then
	Response.Write("<center><a href=""Karakter-Detay/"&CharID2&""" onclick=""pageload('Karakter-Detay/"&CharID2&"');chngtitle('"&CharID2&" > Karakter Detay');return false"" class=""link2"" style=""display:block"">"&CharID2&"&nbsp;&nbsp;&nbsp;")
	If Not onlinechar.eof Then
	If Trim(onlinechar("strcharid"))=CharID2 Then
	Response.Write "<img src=""imgs/on.gif"" align=""absmiddle"" border=""0""></a></center>"
	Else 
	Response.Write "<img src=""imgs/off.gif"" align=""absmiddle"" border=""0""></a></center>"
	End If
	Else 
	Response.Write "<img src=""imgs/off.gif"" align=""absmiddle"" border=""0""></a></center>"
	End If
	End If
	
	If Len(charid3)>0 Then
	Response.Write("<center><a href=""Karakter-Detay/"&CharID3&""" onclick=""pageload('Karakter-Detay/"&CharID3&"');chngtitle('"&CharID3&" > Karakter Detay');return false"" class=""link2"" style=""display:block"">"&CharID3&"&nbsp;&nbsp;&nbsp;")
	If Not onlinechar.eof Then
	If Trim(onlinechar("strcharid"))=CharID3 Then
	Response.Write "<img src=""imgs/on.gif"" align=""absmiddle"" border=""0""></a></center>"
	Else 
	Response.Write "<img src=""imgs/off.gif"" align=""absmiddle"" border=""0""></a></center>"
	End If
	Else 
	Response.Write "<img src=""imgs/off.gif"" align=""absmiddle"" border=""0""></a></center>"
	End If
	End If %>
	<br>
	<%Dim pm, PmKontrol
	Set PmKontrol=Conne.Execute("select count(durum) as toplam from PmBox Where Durum=0 And alici='"&trim(charid1)&"' or Durum=0 And alici='"&trim(charid2)&"' or Durum=0 And alici='"&trim(charid3)&"' ")
	Set pm=Conne.Execute("select count(alici) toplam from pmbox where alici='"&trim(charid1)&"' or alici='"&trim(charid2)&"' or alici='"&trim(charid3)&"' ")
	if Session("yetki")="1" Then%>
	<a href="#" onClick="javascript:pageload('Sayfalar/Gmmenu.asp','1');chngtitle(this.id);return false" id="Gm Menü" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Gm Menü</a><br>
	<% End If %>
	<a href="#" onClick="pageload('Sayfalar/AccountInfo.asp','1');chngtitle(this.id);return false" id="Hesap Bilgileri (MyKOL)" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Hesap Bilgileri (MyKOL)</a><br>
	<a href="#" onClick="pageload('Sayfalar/pmbox.asp','1');chngtitle(this.id);return false" id="Pm Inbox" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Posta Kutusu (<%=pm("toplam")%> / 5)</a><br>
	<a href="#" onClick="pageload('Sayfalar/SellingPanel.Asp','1');chngtitle(this.id);return false" id="Satýþ Paneli" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Satýþ Paneli</a><br>
	<a href="#" onClick="pageload('Sayfalar/debug.asp','1');chngtitle(this.id);return false" id="Askýdan Kurtar" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Askýdan Kurtar</a><br>
	<a href="#" onClick="pageload('Sayfalar/clanleaderchange.asp','1');chngtitle(this.id);return false" id="Clan Devret" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Clan Devret</a><br>
	<a href="#" onClick="pageload('Sayfalar/buycape.asp','1');chngtitle(this.id);return false" id="Pelerin Al" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Clana Pelerin Satýn Al</a><br>
	<a href="#" onClick="pageload('Sayfalar/npdonate.asp','1');chngtitle(this.id);return false" id="Np Baðýþ" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Np Baðýþla</a><br>
	<a href="#" onClick="pageload('Sayfalar/teleportmoradon.asp','1');chngtitle(this.id);return false" id="Teleport To Moradon" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Moradona Iþýnla</a>
	<%If Not PmKontrol.Eof Then
	If PmKontrol("toplam")>0 Then
	Response.Write("<script>alert('"&PmKontrol("toplam")&" Yeni Mesajýnýz Var.\nPosta Kutunuzu Kontrol Ediniz.')</script>")
	End If
	End If
	else Response.Write("Karakteriniz Bulunmuyor.<br />")
	End If %>
	<br />
	<a href="javascript:logout();" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Çýkýþ</a> <br />
	<%Else %>
<style>
.login-text{
background:url("imgs/inputbg.gif") no-repeat ;
border:0;
height:24px;
width:147px;
color:#828282;
font-weight:bold;
text-align:center;
float:left;
font-size:11px;
font-family:Helvetica,Arial,sans-serif;
margin-left:20px;
padding:5px;
}
</style>
<script>
function pwdgoster(){
$('#pwd_hint').css("display","none");
$('#pwd').css("display","block");
document.getElementById('pwd').focus()
}
function pwdgizle(){
if($('#pwd').val()=='')
{
$('#pwd').css("display","none");
$('#pwd_hint').css("display","block");
}
}
</script>
<form action="javascript:logingiris();" method="post" id="loginp" name="loginp">
<input name="username" type="text" class="login-text" id="username" size="20" maxlength="21" value="Kullanýcý Adý" onBlur="if(this.value==''){this.value='Kullanýcý Adý';this.style.color='#828282'}" onFocus="if(this.value=='Kullanýcý Adý'){this.value='';this.style.color='#8E6400'}"><br>
<input name="pwd_hint" type="text" class="login-text" id="pwd_hint" size="20" maxlength="13" value="Þifre" onFocus="pwdgoster()" style="color:#828282"/><br>
<input name="pwd" type="password" class="login-text" id="pwd" size="20" maxlength="13" onBlur="pwdgizle()" style="color:#8E6400;display:none"/>

<div align="center"><input type="submit" value="Giriþ" name="loginb" id="loginb" class="giris"  align="left"></div>

</form>

	<a href="#" onclick="sshowbox();return false" class="link2,hintanchor" onmouseover="showhint('Sanal Klavye', this, event, '100px')"><img src="imgs/keyboard.gif" width="28" height="28" align="left" border="0"></a>
	<a href="#" onclick="javascript:pageload('/Register.html')" class="link2">Kayýt ol</a><br><br>
	<a href="default.asp?cat=sifremi_unuttum" class="link2">Þifremi Unuttum</a>
	<% End If %></td>
	</tr>
        <tr>
          <td height="16" background="imgs/sub_menu_bottom.gif"></td>
	</tr>
	</table>
<%
MenuAyar.Close
Set MenuAyar=Nothing
else
End If%>