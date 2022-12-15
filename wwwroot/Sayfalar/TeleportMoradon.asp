<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<!--#include file="../guvenlik.asp"-->
<%Response.expires=0 
Dim MenuAyar
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='TeleportMoradon'")
If MenuAyar("PSt")=1 Then
if Session("login")="ok" Then
ips=Request.ServerVariables("REMOTE_HOST")%>
<script language="javascript">
function teleportchar(){
$.ajax({
   type: 'POST',
   url: 'sayfalar/teleportmoradon.asp?teleport=ok',
   data: $('#charteleport').serialize() ,
   start:  $('#ortabolum').html('<center><br><br>&nbsp;&nbsp;&nbsp;<img src=imgs/18-1.gif><br>Teleporting. Please Wait...</center>'),

success: function(ajaxCevap) {
      $('#ortabolum').html(ajaxCevap);
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
</style><br><img src="imgs/teleporttomoradon.gif"><br /><br /><br />
<%
set chars=Conne.Execute("select * from account_char where straccountid='"&Session("username")&"'")
char1=trim(chars("strcharid1"))
char2=trim(chars("strcharid2"))
char3=trim(chars("strcharid3"))
teleport=secur(Request.Querystring("teleport"))
if teleport="" Then%>
<form action="javascript:teleportchar();" method="post" id="charteleport" name="charteleport">
<table width="225" border="0">
<tr><td> </td></tr>
  <tr>
    <td class="txt"><b>Karakter seçiniz :</b> <select name="char" class="inpt" style="padding-top:2px">
<%if len(char1)>0 Then
Response.Write "<option value="&char1&" style=""height:15px"">"&char1&"</option>"
End If
if len(char2)>0 Then
Response.Write "<option value="&char2&" style=""height:15px"">"&char2&"</option>"
End If
if len(char3)>0 Then
Response.Write "<option value="&char3&" style=""height:15px"">"&char3&"</option>"
End If%>
    </select></td>
  </tr>
  <tr>
    <td align="center"><input name="submit" type="submit" value="Iþýnla" style="color:#8E6400;font-weight:bold;font-size:10px;" class="styleform" /></td>
    </tr>
</table>

  <br />
</form>
<%elseif teleport="ok" Then
char=secur(request.form("char"))
if char=char1 or char=char2 or char=char3 Then
set charsearch=Conne.Execute("select * from account_char where straccountid='"&Session("username")&"' and strcharid1='"&char&"' or strcharid2='"&char&"' or strcharid3='"&char&"'")
if not chars.eof Then
set isinla=Conne.Execute("update userdata set zone='21', PX='31200', PZ='40200', PY='0' where struserid='"&char&"'")
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&char&" Karakteri Moradona ýþýnlandý.','"&now&"')")
Response.Write "<b><br>Karakteriniz Moradona ýþýnlandý.</b><script>cal('warp_act_0.mp3')</script>"
Else
Response.Write "Karakter bulunamadý.."
End If
Else
Response.Write "Karakter bulunamadý.."
End If
End If
End If
Else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>
