<style>
td {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:12px;
}
</style><%if Session("sqllogin")="" Then
serverip=request.cookies("remember")("serverip")
dbname=request.cookies("remember")("dbname")
loginid=request.cookies("remember")("logid")
loginpwd=request.cookies("remember")("pwd")
giris=request.cookies("remember")("giris")
If loginid<>"" And dbname<>"" And serverip<>"" And loginpwd<>"" And giris="ok" Then
Session("sqllogin")="ok"
Session("connect")=array(serverip,dbname,loginid,loginpwd)
Response.Redirect("default.asp")
End If
%>
<center><form action="default.asp?pro=login" method="post">
<table style="position:absolute;top:30%;left:35%">
<tr><td colspan="2" align="center">Sql Table Editor</td></tr>
<tr>
<td>Server IP:</td>
<td><input type="text" name="serverip" value="<%=serverip%>"></td>
</tr>
<tr>
<td>Database Name:</td>
<td><input type="text" name="dbname" value="<%=dbname%>"></td>
</tr>
<tr>
<td>User ID:</td>
<td><input type="text" name="loginid" value="<%=loginid%>"></td>
</tr>
<tr>
<td>User Pwd:</td>
<td><input type="text" name="loginpwd" value="<%=loginpwd%>"></td></tr>
<tr>
<td>Ayarlarý Anýmsa</td>
<td><input type="checkbox" name="remember" value="check" <%if serverip<>"" Then Response.Write "checked"%>></td></tr>
<tr>
<td colspan="2" align="center"><input type="submit"  value="Giriþ Yap"></td>
</tr></table></form>
</center>
<%else
connect=Session("connect")


Sunucu = connect(0)
VeriTabani = connect(1)
Kullanici = connect(2)
Sifre = connect(3)

Set conne = Server.CreateObject("ADODB.Connection")
conne.open= "driver={SQL Server};server="&sunucu&";database="&veritabani&";uid="&kullanici&";pwd="&sifre&"" 
%>
<script type="text/javascript" src="instantedit.js"></script>
<script>
function rowsil(tablename,rowid){
document.getElementById('row'+rowid).style.display='none';

remotos = new datosServidor;
nt = remotos.enviar("anindaeditle.asp?fieldname="+encodeURI(tablename)+"&content="+encodeURI(rowid));
}
</script>
<style>
a{
color:#003A7A;
font-size:12px;
text-decoration:none
}
body,th {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:12px;
}
td {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:11px;
}

</style>
Database: <%Response.Write veritabani&" &nbsp;<a href=""default.asp?pro=logout"">Çýkýþ</a><br>"
topen=Request.Querystring("topen")
if topen="" Then
topen="list"
elseif topen<>"list" Then
tablename=Request.Querystring("tablename")
End If
if topen="list" Then
set tables=Conne.Execute("select name,id from sysobjects where xtype='U' order by name")
Response.Write "<div style=""overflow:auto;width:400;height:90%"">"
do while not tables.eof
Response.Write "<a href=""default.asp?topen=table&tablename="&tables(0)&""">"&tables(0)&"<br>"&vbcrlf
tables.movenext
loop
Response.Write "</div>"
elseif topen="truncatetable" Then
tablename=Request.Querystring("tablename")
Conne.Execute("Truncate Table "&tablename)
Response.Redirect("default.asp?topen=table&tablename="&tablename)
elseif topen="table" Then
set tables=Conne.Execute("select id from sysobjects where name='"&tablename&"' and xtype='U'")
if tables.eof Then
Response.Write "Tablo Bulunamadý"
Response.End
End If
set sutunlar=Conne.Execute("select name,length,usertype,isnullable from syscolumns where id="&tables(0)&" order by colorder ")
if not sutunlar.eof Then
set icerik=Conne.Execute("select * from "&tablename&"")
set iceriktop=Conne.Execute("select count(*) as toplam from "&tablename&"")
Response.Write "<br><a href=""default.asp?topen=truncatetable&tablename="&tablename&""" onclick=""return confirm('Tablodaki Bütün Kayýtlar Silinecek.\n Bunu Yapmak Istediðinizden Eminmisiniz?')"">Tabloyu Boþalt</a><br><a href=""default.asp"">< Listeye Dön</a>"
if not icerik.eof Then
sayfa=Request.Querystring("sayfa")
gosterim=100
if sayfa="" Then
sayfa=1
End If
icerik.move((sayfa-1)*gosterim)
Response.Write "<br>Sayfalar:"
tsayfa=int(iceriktop("toplam")/gosterim)
if iceriktop("toplam") mod gosterim>0 Then tsayfa=tsayfa+1
for x=1 to tsayfa
if cint(sayfa)=x Then
Response.Write "<a href=""default.asp?topen="&topen&"&tablename="&tablename&"&sayfa="&x&"""><b>"&x&"</b></a> "&vbcrlf
else
Response.Write "<a href=""default.asp?topen="&topen&"&tablename="&tablename&"&sayfa="&x&""">"&x&"</a> "&vbcrlf
End If
next%>
<div style="overflow:scroll;width:100%;height:100%">
<table cellspacing="1" border="1">
<tr style="background-color:#D4D0C8;font-family:arial;font-size:12px;"><td>&nbsp;</td>
<%stn=0
dim sutun(100)
dim sutunnull(100)
do while not sutunlar.eof
Response.Write "<td>"&sutunlar(0)&"</td>"
sutun(stn)=sutunlar(2)
sutunnull(stn)=sutunlar("isnullable")
sutunlar.movenext
stn=stn+1
loop

rowid=0
hno=1
for sayfala=1 to gosterim
if icerik.eof Then exit for
Response.Write "<tr style=""font-family:arial;font-size:12px;"" id=""row"&((sayfa-1)*gosterim)+rowid&""" ><td style=""background-color:#D4D0C8;color:#000"">"&(sayfa-1)*gosterim+rowid+1&"</td>"
for x=0 to stn-1
if sutun(x)=20 Then
Response.Write "<td><span id="""&tablename&","&((sayfa-1)*gosterim)+rowid&","&x&""" style=""display:block"">&lt;IMAGE&gt;"
else
Response.Write "<td><span id="""&tablename&","&((sayfa-1)*gosterim)+rowid&","&x&","&sutunnull(x)&""" class=""editText"" style=""display:block"">"
if isnull(icerik(x))=true Then
Response.Write "NULL"
elseif icerik(x)="" Then
Response.Write "&nbsp;"
else
Response.Write trim(icerik(x))
End If
End If

Response.Write "</span></td>"
hno=hno+1
next

Response.Write "<td><a href=""anindaeditle.asp?islem=sil&fieldname="&tablename&","&rowid&""">Sil</a></td></tr>"&vbcrlf
icerik.movenext
rowid=rowid+1
next
else
Response.Write "<br><br>Kayýt Bulunamadý!"
End If

End If
%>
</tr>
</table>
<%End If
End If
if Request.Querystring("pro")="login" Then
serverip=request.form("serverip")
dbname=request.form("dbname")
logid=request.form("loginid")
pwd=request.form("loginpwd")
if request.form("remember")="check" Then
Response.Cookies("remember")("serverip")=serverip
Response.Cookies("remember")("dbname")=dbname
Response.Cookies("remember")("logid")=logid
Response.Cookies("remember")("pwd")=pwd
Response.Cookies("remember")("giris")="ok"
Response.Cookies("remember").expires=now+999
else
Response.Cookies("remember")("serverip")=""
Response.Cookies("remember")("dbname")=""
Response.Cookies("remember")("logid")=""
Response.Cookies("remember")("pwd")=""
End If
Session("sqllogin")="ok"
Session("connect")=array(serverip,dbname,logid,pwd)
Response.Redirect("default.asp")
elseif Request.Querystring("pro")="logout" Then
Response.Cookies("remember")("giris")=""
Session.abandon
Response.Redirect("default.asp")
End If
%></div>