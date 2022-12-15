<!--#include file="../_inc/conn.asp"-->
<!--#include file="../Function.asp"-->
<style>
td {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:12px;
}
</style>
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
<%topen=Request.Querystring("topen")
tablename=Request.Querystring("tablename")
if topen="truncatetable" Then
tablename=Request.Querystring("tablename")
Conne.Execute("Truncate Table "&tablename)
Response.Redirect("TableManager.asp?topen=table&tablename="&tablename)
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
Response.Write "<a href=""TableManager.asp?topen=truncatetable&tablename="&tablename&""" onclick=""return confirm('Tablodaki Bütün Kayýtlar Silinecek.\n Bunu Yapmak Istediðinizden Eminmisiniz?')"">Tabloyu Boþalt</a>"
if not icerik.eof Then
sayfa=Request.Querystring("sayfa")
gosterim=100
if sayfa="" Then
sayfa=1
End If
icerik.move((sayfa-1)*gosterim)
Response.Write "<br><form action=""TableManager.asp?topen="&topen&"&tablename="&tablename&""" method=""get"">Sayfalar: <select >"
tsayfa=int(iceriktop("toplam")/gosterim)
if iceriktop("toplam") mod gosterim>0 Then tsayfa=tsayfa+1
For x=1 To tsayfa
If cint(sayfa)=x Then
Response.Write "<option value="""&x&""" onclick=""this.form.action=this.form.action+'&sayfa='+this.value"">"&x&"</option>"&vbcrlf
Else
Response.Write "<option value="""&x&""" onclick=""this.form.action=this.form.action+'&sayfa='+this.value"">"&x&"</option>"&vbcrlf
End If
next
Response.Write "</select><input type=""submit""></form>"%>
<div style="overflow:scroll;width:100%;height:80%">
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

%></div>