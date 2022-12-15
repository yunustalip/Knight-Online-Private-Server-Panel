<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<%Response.expires=0 
Dim MenuAyar,ksira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Search'")
If MenuAyar("PSt")=1 Then%>
<meta http-equiv="content-type" content="text/html; charset=windows-1254">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" >
<style type="text/css">
<!--
.style21 {color: #FFFFFF; font-weight:bold}
-->
</style>
  <script language="javascript">
function ara(){
$.ajax({
   type: 'get',
   url: 'search.asp',
   data: $('#arama').serialize() ,
   start:  $('#ortabolum').html('&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif><br>Arama yapılıyor Lütfen Bekleyin.'),
   success: function(ajaxCevap) {
      $('#ortabolum').html(ajaxCevap);
   }
});
}
</script>
<%Response.Charset = "iso-8859-9"
Dim ara
ara=secur(Request.Querystring("ara"))
if ara="" Then%><br><img src="imgs/aramamotoru.gif" /><br><br><br>
<form action="javascript:ara()" method="get" id="arama" name="arama">
<table width="270" border="0">
  <tr>
    <td width="91">Ara: </td>
    <td width="169"><select name="ara" style="background:#8E6400;color:white;font-weight:bold;font-size:10px;font-family:verdana">
	<option value="user">User</option>
	<option value="clan">Clan</option>
	</select></td>
  </tr>
  <tr>
    <td>User / Clan Adı: </td>
    <td><input type="text" name="isim" style="background: #8E6400; color: white;font-weight:bold"/></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input type="submit" value="Ara" /></td>
    </tr>
</table>

</form>
<%elseif ara="user" Then
Dim isim
Dim user
isim=trim(secur(Request.Querystring("isim")))
if len(isim)>1 Then

set user=Conne.Execute("select struserid,loyalty,class from userdata where struserid='"&isim&"' ")
if not user.eof Then%><br />
<table width="300">
<tr width="300">
<td background="imgs/menubg.gif" class="style21" width="100">Kullanıcı Adı</td>
<td background="imgs/menubg.gif" class="style21" width="100">National Point</td>
<td background="imgs/menubg.gif" class="style21" width="100">Tür</td>
</tr>
<%do while not user.eof%>
<tr width="300">
<td width="100"><a href="Karakter-Detay/<%=trim(user("struserid"))%>" onclick="pageload('Karakter-Detay/<%=trim(user("struserid"))%>');return false"><%=user("struserid")%></a></td>
<td width="100"><%=user("loyalty")%></td>
<td width="100"><%Response.Write cla(user("class"))%></td>
</tr>
<%user.movenext
loop
Response.Write "</table>"
else
Response.Write "<br><br><b>Kullanıcı Bulunamadı.</b>"
End If
else
Response.Write "<br><br><b>En Az 2 karakter yazmanız gerekmektedir.</b>"
End If
elseif ara="clan" Then
isim=trim(secur(Request.Querystring("isim")))
if len(isim)>1 Then
Dim clan
set clan=Conne.Execute("select idnum,idname,points from knights where idname='"&isim&"'")
if not clan.eof Then%>
<table width="200">
<tr >
<td background="imgs/menubg.gif" class="style21">Clan Adı</td>
<td background="imgs/menubg.gif" class="style21">National Point</td>

</tr>
<%do while not clan.eof%>

<tr>
<td><a href="Clan-Detay/<%=trim(clan("idname"))&","&clan("idnum")%>" onclick="pageload('Clan-Detay/<%=trim(clan("idname"))&","&clan("idnum")%>');return false"><%=clan("idname")%></a></td>
<td><%=clan("points")%></a></td>

</tr>
<%clan.movenext
loop
Response.Write "</table>"
else
Response.Write "<br><br><b>Kullanıcı Bulunamadı</b>"
End If
else
Response.Write "<br><br><b>En Az 2 karakter yazmanız gerekmektedir.</b>"
End If
else
End If

else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafından kapatılmıştır.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing
%>