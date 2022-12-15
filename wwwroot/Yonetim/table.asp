<%Sunucu = "localhost"
VeriTabani = "kn_online"
Kullanici = "talip"
Sifre = "864327"



Set conne = Server.CreateObject("ADODB.Connection")
conne.open= "driver={SQL Server};server="&sunucu&";database="&veritabani&";uid="&kullanici&";pwd="&sifre&"" 
%>
<script type="text/javascript" src="instantedit.js"></script>
<script>
function rowsil(tablename,rowid){
document.getElementById('row'+rowid).style.display='none';

remotos = new datosServidor;
nt = remotos.enviar("anindaeditle.asp?fieldname="+encodeURI(tablename)+"&content="+encodeURI(rowid));

</script>
<style>
a{
color:#003A7A;
font-size:14px
}
</style>
<div style="overflow:auto;width:400%"><% topen=Request.Querystring("topen")
if topen="" Then
topen="list"
elseif topen<>"list" Then
tablename=Request.Querystring("tablename")
End If
if topen="list" Then
set tables=Conne.Execute("select name,id from sysobjects where xtype='U' order by name")
do while not tables.eof
Response.Write "<a href=""?topen=table&tablename="&tables(0)&""">"&tables(0)&"<br>"
tables.movenext
loop
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
if not icerik.eof Then
sayfa=Request.Querystring("sayfa")
gosterim=100
if sayfa="" Then
sayfa=1
End If
icerik.move((sayfa-1)*gosterim)
Response.Write "<a href=""table.asp"">< Listeye Dön</a><br>Sayfalar:"
tsayfa=int(iceriktop("toplam")/gosterim)
if iceriktop("toplam") mod gosterim>0 Then tsayfa=tsayfa+1
for x=1 to tsayfa
if cint(sayfa)=x Then
Response.Write "<a href=""?topen="&topen&"&tablename="&tablename&"&sayfa="&x&"""><b>"&x&"</b></a> "
else
Response.Write "<a href=""?topen="&topen&"&tablename="&tablename&"&sayfa="&x&""">"&x&"</a> "
End If
next%>
<table cellspacing="1" border="1">
<tr style="background-color:#D4D0C8;font-family:arial;font-size:12px;"><td>&nbsp;</td>
<%stn=0
dim sutun(100)
do while not sutunlar.eof
Response.Write "<td>"&sutunlar(0)&"</td>"
sutun(stn)=sutunlar(2)
sutunlar.movenext
stn=stn+1
loop

rowid=0
hno=1
for sayfala=1 to gosterim
if icerik.eof Then exit for
Response.Write "<tr style=""font-family:arial;font-size:12px;"" id=""row"&((sayfa-1)*gosterim)+rowid&""" ><td>"&(sayfa-1)*gosterim+rowid+1&"</td>"
for x=0 to stn-1
if sutun(x)=20 Then
Response.Write "<td><span id="""&tablename&","&((sayfa-1)*gosterim)+rowid&","&x&""" style=""display:block"">&lt;IMAGE&gt;"
else
Response.Write "<td><span id="""&tablename&","&((sayfa-1)*gosterim)+rowid&","&x&""" class=""editText"" style=""display:block"">"
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

End If
else
Response.Write "Kayýt Bulunamadý!"
End If
%>
</tr>
<%End If
%></div>