<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="Function.asp"-->
<%
REFERER_URL = Request.ServerVariables("HTTP_REFERER")

If InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
Else
REFERER_DOMAIN = Left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If


If REFERER_DOMAIN="http://"&Request.ServerVariables("server_name") or  REFERER_DOMAIN="http://www."&Request.ServerVariables("server_name") Then
Else
yn("/Kim-Kimi-Kesmis")
End If
%><script>
function kimkimikesti(url,forma){
$.ajax({
   type: 'GET',
   url: url,
   data: $('#'+forma).serialize() ,
   start:  $('#cevap').html('<center>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif><br>Arama Yapýlýyor. Lütfen Bekleyin...</center>'),
   success: function(ajaxCevap) {
      $('#cevap').html(ajaxCevap);
   }
});
$('#gun').val()
}
function kimkimikesti2(url){
$.ajax({
   type: 'GET',
   url: url,
   start:  $('#cevap').html('<center>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif><br>Arama Yapýlýyor. Lütfen Bekleyin...</center>'),
   success: function(ajaxCevap) {
      $('#cevap').html(ajaxCevap);
   }
});
$('#gun').val()
}
</script>
<%response.charset="iso-8859-9"
Dim MenuAyar,ksira,gun,ay,yil,logadresi,sid,nesne
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='KimKimKesmis'")
If MenuAyar("PSt")=1 Then
gun=Request.Querystring("gun")
ay=Request.Querystring("ay")
yil=year(now)
if Request.Querystring("sayfa")="" Then%><br />
<center>
<img src="imgs/kimkimikesmis.gif" /></center><br /><br />
<br>
<form id="form1" name="form1" method="get" action="kimkimikesmis.asp" onsubmit="javascript:kimkimikesti('kimkimikesmis.asp?sayfa=1','form1');return false">
  <div align="center">
    <label>
      Char Adý:
      <input style="background: #8E6400; color: white" type="text" name="sid" >
    </label>
    <label>
    <label>
      Gün:
      <input style="background: #8E6400; color: white" name="gun" id="gun" type="text" value="<%=server.htmlencode(day(now))%>" size="5" />
    </label><input type="hidden" name="sayfa" value="1">
  Ay:
  <input style="background: #8E6400; color: white" name="ay" id="ay" type="text" value="<%=server.htmlencode(month(now))%>" size="10" />
  <label>
      <input style="background: #8E6400; color: white; font-weight:bold" type="submit" value="Göster" />
    </label>
  </div>
</form>
 <form id="form2" name="form2" method="get" action="kimkimikesmis.asp" onsubmit="javascript:kimkimikesti('kimkimikesmis.asp?sayfa=1','form2');return false">
  <div align="center">
    <label>
      Gün:
      <input style="background: #8E6400; color: white" name="gun" id="gun" type="text" value="<%=server.htmlencode(day(now))%>" gun" size="10" />
    </label><input type="hidden" value="1" name="sayfa">
  Ay:
  <input style="background: #8E6400; color: white" name="ay" id="ay" type="text" value="<%=server.htmlencode(month(now))%>" size="10" />
  <label>
     <input style="background: #8E6400; color: white;font-weight:bold" type="submit" value="Göster" />
  </label>
  </div>
</form>


<div id="cevap">
<%else
if isnumeric(gun)=false or isnumeric(ay)=false or len(gun)>2 or len(ay)>2 Then
Response.Write "Lütfen Sadece Sayýsal Deðer Giriniz..."
Response.End
End If
logadresi="D:\KO\SERVER FILES\3 - Ebenezer\DeathLog-" 'Ebenezerin klasöründe saklanan loglara göre yolu düzenleyiniz.

sid=lcase(Request.Querystring("sid"))

if gun<>"" and ay<>"" Then

set nesne = Server.CreateObject("Scripting.FileSystemObject")
if nesne.FileExists(logadresi&yil&"-"&ay&"-"&gun&".txt")=True Then

set qa = nesne.OpenTextFile(logadresi&yil&"-"&ay&"-"&gun&".txt",1,0)%>
 <center>Tarih: <%if gun<>"" and ay<>"" Then
Response.Write gun&"."&ay&"."&yil
else
Response.Write day(now)&"."&month(now)&"."&yil
End If

Response.Write "<table width=""510"" border=""1"" cellspacing=""1"" cellpadding=""1"" bordercolor=""#8E6400"">"&_
"    <tr>"&_
"      <td><CENTER>Kesen</td>"&_
"      <td><CENTER>Kesilen</td>"&_
"      <td><CENTER>Önceki NP</td>"&_
"      <td><CENTER>Sonraki NP</td>"&_
"      <td><CENTER>Harita</td>"&_
"      <td><CENTER>Saat</td>"&_
"    </tr>"
Do While Not qa.AtEndOfStream

Satir = qa.ReadLine
parca=split(satir,",")

if len(Request.Querystring("sid"))>0 Then
if sid=lcase(parca(4)) Then
if parca(5)>0 and parca(13)>0 Then
If Trim(parca(3))="21" Then 
zone="Moradon"
elseif Trim(parca(3))="1" Then 
zone="Luferson Castle"
elseif Trim(parca(3))="2" Then 
zone="Elmorad Castle"
elseif Trim(parca(3))="201" Then 
zone="Colony Zone"
elseif Trim(parca(3))="202" Then 
zone="Ardream"
elseif Trim(parca(3))="30" Then 
zone="Delos"
elseif Trim(parca(3))="48" Then 
zone="Arena"
elseif Trim(parca(3))="101" Then 
zone="Lunar War"
elseif Trim(parca(3))="102" Then 
zone="Dark Lunar War"
elseif Trim(parca(3))="103" or Trim(parca(3))="111" Then 
zone="War Zone"
elseif Trim(parca(3))="11" Then 
zone="Karus Eslant"
elseif Trim(parca(3))="12" Then 
zone="El Morad Eslant"
elseif Trim(parca(3))="31" Then 
zone="Bi-Frost"
elseif Trim(parca(3))="51" or Trim(parca(3))="52" or Trim(parca(3))="53" or Trim(parca(3))="54" or Trim(parca(3))="55" Then 
zone="Forgetten Temple Zone"
elseif Trim(parca(3))="32" Then 
zone="Hell Abyss"
elseif Trim(parca(3))="33" Then 
zone="Isiloon Floor"
Else
set zoneid=Conne.Execute("select bz from zone_info where zoneno="&Trim(parca(3)))
if not zoneid.eof Then
zone=zoneid("bz")
else
zone="-"
End If
End If
%>
 <tr >
   <td align="center"><a href="#" onclick="javascript:kimkimikesti2('kimkimikesmis.asp?sayfa=1&<%Response.Write "gun="&gun&"&ay="&ay&"&sid="&parca(4)&" ');return false "">"&parca(4)%></td>
   <td align="center"><%=parca(12)%></td>
   <td align="center"><%=parca(7)%></td>
   <td align="center"><%=parca(10)%></td>
   <td align="center"><%=zone%></td>
   <td align="center"><%=parca(0)&":"&parca(1)&":"&parca(2)%></td>
 </tr>
<%
End If
End If
else
If parca(5)>0 and parca(13)>0 Then
If Trim(parca(3))="21" Then 
zone="Moradon"
elseif Trim(parca(3))="1" Then 
zone="Luferson Castle"
elseif Trim(parca(3))="2" Then 
zone="Elmorad Castle"
elseif Trim(parca(3))="201" Then 
zone="Colony Zone"
elseif Trim(parca(3))="202" Then 
zone="Ardream"
elseif Trim(parca(3))="30" Then 
zone="Delos"
elseif Trim(parca(3))="48" Then 
zone="Arena"
elseif Trim(parca(3))="101" Then 
zone="Lunar War"
elseif Trim(parca(3))="102" Then 
zone="Dark Lunar War"
elseif Trim(parca(3))="103" or Trim(parca(3))="111" Then 
zone="War Zone"
elseif Trim(parca(3))="11" Then 
zone="Karus Eslant"
elseif Trim(parca(3))="12" Then 
zone="El Morad Eslant"
elseif Trim(parca(3))="31" Then 
zone="Bi-Frost"
elseif Trim(parca(3))="51" or Trim(parca(3))="52" or Trim(parca(3))="53" or Trim(parca(3))="54" or Trim(parca(3))="55" Then 
zone="Forgetten Temple Zone"
elseif Trim(parca(3))="32" Then 
zone="Hell Abyss"
elseif Trim(parca(3))="33" Then 
zone="Isiloon Floor"
Else
set zoneid=Conne.Execute("select bz from zone_info where zoneno="&Trim(parca(3)))
if not zoneid.eof Then
zone=zoneid("bz")
else
zone="-"
End If
End If
%>
 <tr>
   <td align="center"><a href="#" onclick="javascript:kimkimikesti2('kimkimikesmis.asp?sayfa=1&<%Response.Write "gun="&gun&"&ay="&ay&"&sid="&parca(4)&" ');return false "">"&parca(4)%></td>
   <td align="center"><%=parca(12)%></td>
   <td align="center"><%=parca(7)%></td>
   <td align="center"><%=parca(10)%></td>
   <td align="center"><%=zone%></td>
   <td align="center"><%=parca(0)&":"&parca(1)&":"&parca(2)%></td>
 </tr>
<%
End If
End If

Loop
Else
Response.Write "<tr><td align=""center"" colspan=""5""><b>Bu Tarihe Ait Log Bulunamadý. <br>Tarih: "&gun&"."&ay&"."&yil&"</td></tr>"
End If
Response.Write "</table>"
End If
%></div>

</div>
<%End If

Else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If

MenuAyar.Close
Set MenuAyar=Nothing%>