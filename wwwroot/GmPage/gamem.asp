<!--#include file="../_inc/DbSettings.Asp"-->
<!--#include file="../_inc/Connect.Asp"-->
<%if Session("yetki")="1" Then
Dim Bag,Conne
Set Bag = New Baglanti
Set Conne = Bag.Connect(Sunucu,VeriTabani,Kullanici,Sifre)
%>
<!--#include file="../function.asp"-->
<!--#include file="../md5.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9">
<base href="http://<%=Request.ServerVariables("server_name")%>/">
<link href="../css/webstyle.css" rel="stylesheet" type="text/css">
<style>
td {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:11px;
}

</style>
<body background="imgs/bgbr.jpg">
<%Response.Expires=0
dim user,nick,username,pwd,gmkontrol,gmkont,gmpwd1,gmpwd2
user=secur(Request.Querystring("user"))
nick=secur(Request.Querystring("nick"))

if user="gmkomut" Then
dim komut,knick,userkont
Function ConvertFromUTF8(sIn)
Dim oIn: Set oIn = CreateObject("ADODB.Stream")

oIn.Open
oIn.CharSet = "WIndows-1254"
oIn.WriteText sIn
oIn.Position = 0
oIn.CharSet = "UTF-8"
ConvertFromUTF8 = oIn.ReadText
oIn.Close

End Function

komut=ConvertFromUTF8(Request.Querystring("komut"))
if instr(komut,"/kill")>0 Then
knick=mid(komut,instr(komut,"/kill")+6,len(komut)-5)
set userkont=Conne.Execute("select authority from userdata where struserid='"&knick&"' ")
if userkont.Eof Then
Response.Write("Bu Nickli Bir Karakter bulunamadı.")
Response.End
End If
if userkont("authority")="0" Then
Response.Write "Bu Kullanıcı Oyun Yöneticisidir. Oyundan Çıkaramazsınız."
Response.End
else
Conne.Execute("delete currentuser where strcharid='"&knick&"'")
End If
End If
if komut="" Then
komut=" "
End If
dim komutx,komuttara
if left(komut,1)="/" Then
if instr(komut," ")>0 Then
komutx=left(komut,instr(komut," ")-1)
else
komutx=komut
End If
set komuttara=Conne.Execute("select gmyetki from siteayar")
dim parca,kmtvar,kmtsy
kmtvar="0"
parca=split(komuttara("gmyetki"),vbcrlf)
For kmtsy=0 To ubound(parca)
If lcase(trim(komutx))=lcase(trim(parca(kmtsy))) Then
kmtvar="1"
Exit For
End If
Next
if kmtvar="0" Then
Response.Write("<script>alert('Bu Komutu Kullanma Yetkiniz Yok')</script>")
Response.End
End If
End If
Dim Shell
Set Shell = Server.CreateObject("WScript.Shell")
Shell.Run(server.mappath("cmdEb.exe")&" "&komut)

Elseif user="ban" Then
dim onl
set onl=Conne.Execute("select * from currentuser where strCharID='"&nick&"'")
if not onl.eof Then%>
<form action="GmPage/gamem.asp?user=banla&nick=<%=nick%>" method="post">
<%else%>
<form action="GmPage/gamem.asp?user=banla&nick=<%=nick%>" method="post">
<%End If%><table>
  <tr>
    <td ><b>Banlanılacak Nick </b>:</td>
    <td ><input name="charid" type="text" disabled value="<%=nick%>"/></td>
  </tr>
  <tr>
    <td><b>Banlanma Sebebi :</b></td>
    <td><input type="text" name="sebep" size="25"/></td>
  </tr>
  <tr>
    <td><b>Banlanılacak gün sayısı:</b></td>
    <td><input type="text" name="gun" size="4" maxlength="4"/></td>
  </tr>
  <tr>
    <td colspan="2"><input type="submit" value="Banla" style="width:250px;height:25px;font-weight:bold;cursor:pointer" class="styleform"/><%if not onl.eof Then
Response.Write "<br><input type=""button"" value="""&nick&" Nickli Oyuncuyu Disconnect Et"" onClick=""location.href='GmPage/gamem.asp?user=gmkomut&komut=/kill "&nick&"'"" style=""width:250px;height:25px;font-weight:bold;cursor:pointer"" class=""styleform""/>"%><br>
<input type="button" value="Giriş Yapılan Ip Adresini Banla (<%=onl("strclientip")%>)" onClick="location.href='GmPage/Gamem.asp?user=ipban&ip=<%=onl("strclientip")%>'" style="width:250px;height:25px;font-weight:bold;cursor:pointer" class="styleform"/>
	</td>
  </tr>
<%End If%>
</table>

</form>

<br><br>
<%
elseif user="banla" Then
dim sebep,gun
nick=secur(Request.Querystring("nick"))
sebep=secur(request.form("sebep"))
gun=secur(request.form("gun"))

if nick="" or sebep="" or gun="" Then
Response.Write "Boş alan bırakmayın."
Response.End
else

if not isnumeric(gun)=True Then
Response.Write "Gün bölümüne sadece sayı yazın."
Response.End
End If
dim userban
set onl=Conne.Execute("select * from currentuser where strCharID='"&nick&"'")
Set userban = Conne.Execute("Select * From userdata where struserid='"&nick&"'")

if not userban.eof Then
if userban("authority")<>"0" Then
dim ban
if onl.eof Then

Set ban = Server.CreateObject("ADODB.Recordset")
sql = "Select authority,yasaksebep,yasakgun,bancount From userdata where struserid='"&nick&"'"
ban.open sql,conne,1,3
ban("bancount")=ban("bancount")+1
ban("yasaksebep")=sebep
ban("yasakgun")=dateadd("d",gun,now)
ban("authority")=255
ban.update
ban.close
set ban=nothing
Response.Write "<b>"&nick&"&nbsp; Is Banned</b>"
Response.Write "<iframe src=""GmPage/Gamem.asp?user=gmkomut&komut="&nick&" Is Banned For: "&sebep&" "&gun&" Days"" style=""display:none""></iframe>"
application("notice")=nick&" Is Banned For: "&sebep&" "&gun&" Days|"&now()
application("noticeuser")=""
else
Response.Write "<iframe src=""GmPage/Gamem.asp?user=gmkomut&komut=/kill "&nick&""" style=""display:none""></iframe>"
Conne.Execute("delete currentuser where strcharid='"&nick&"'")
Set ban = Server.CreateObject("ADODB.Recordset")
sql = "Select authority,yasaksebep,yasakgun,bancount From userdata where struserid='"&nick&"'"
ban.open sql,conne,1,3
ban("bancount")=ban("bancount")+1
ban("yasaksebep")=sebep
ban("yasakgun")=dateadd("d",gun,now)
ban("authority")=255
ban.update
ban.close
set ban=nothing

Response.Write "<iframe src=""GmPage/Gamem.asp?user=gmkomut&komut="&nick&" Is Banned For: "&sebep&" "&gun&" Days"" style=""display:none""></iframe>"
application("notice")=nick&" Is Banned For: "&sebep&" "&gun&" Days|"&now()
application("noticeuser")=""
Response.Write "<b>"&nick&"&nbsp; Is Banned</b>"
End If
else
Response.Write "Game Master lar sadece Admin Yönetim panelinden banlanabilirler."
End If
else
Response.Write "Karakter  Bulunamadı."
End If


End If

elseif user="dc" Then

nick=secur(Request.Querystring("nick"))
Response.Write nick&"&nbsp; Oyundan Atılsınmı (Disconnect) ?<br><a href=""gamem.asp?user=gmkomut"&"&komut=/kill "&nick&""">Evet</a>&nbsp;&nbsp;<a href=javascript:window.close();>Hayır</a>"
elseif user="bankaldir" Then

nick=secur(Request.Querystring("nick"))
Response.Write nick&"&nbsp;karakterinin yasağını kaldırmak istiyomusunuz ?<br><a href=""GmPage/gamem.asp?user=unban&nick="&nick&""">Evet</a>&nbsp;&nbsp;<a href=javascript:window.close();>Hayır</a>"
elseif user="unban" Then
nick=secur(Request.Querystring("nick"))

Set unban = Conne.Execute("Select * From userdata where struserid='"&nick&"'")

if not unban.eof Then
if unban("authority")="0" Then
Response.Write "Game Master işlemleri sadece admin panelinden gerçekleştirilir."
else
Conne.Execute("update userdata set authority=1,yasaksebep=null,yasakgun=null where struserid='"&nick&"'")
Response.Write Nick&"&nbsp; karakterinin yasağı kaldırılmıştır."
End If

Else
Response.Write "Karakter Bulunamadı."
End If
Elseif user="ipban" Then
ip=Request.Querystring("ip")
Conne.Execute("insert into bannedip values('"&ip&"')")
Response.Write "<br><b>"&ip&" No lu Ip Adresi Banlanmıştır."
ElseIf user="runprograms" Then
komut=Request.Querystring("komut")
Set Shell = Server.CreateObject("WScript.Shell")
Shell.Exec(komut)
End If
End If
%>