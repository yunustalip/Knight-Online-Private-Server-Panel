<%Session.CodePage=1254
Response.Charset = "windows-1254"
on error resume next
connect=Session("connect")
Sunucu = connect(0)
VeriTabani = connect(1)
Kullanici = connect(2)
Sifre = connect(3)
Set conne = Server.CreateObject("ADODB.Connection")
conne.open= "driver={SQL Server};server="&sunucu&";database="&veritabani&";uid="&kullanici&";pwd="&sifre&"" 

Response.Buffer = True
Response.contenttype = "text/html; charset=windows-1254"
YaziYazilanAlan = Request("fieldname")
icerik = Request("content")
islem= Request("islem")
if islem="sil" Then
parcala=split(YaziYazilanAlan,",")
table=parcala(0)
rownum=parcala(1)
Set userlogin = Server.CreateObject("ADODB.Recordset")
strSQL ="SELECT * FROM "&table
userlogin.open strSQL,Conne,1,3
userlogin.move(rownum)
userlogin.delete
userlogin.close
set userlogin=nothing
Response.Redirect("default.asp?tablename="&table&"&topen=table")
else
parcala=split(YaziYazilanAlan,",")
table=parcala(0)
rownum=parcala(1)
colnum=cint(parcala(2))
nable=cint(parcala(3))

if icerik="" and nable=0 or icerik=" " and nable=0 or icerik=Null and nable=0 or icerik="NULL" and nable=0 or icerik="Null" and nable=0 Then
Response.Write "Bu Alan Boþ Býrakýlamaz !"
Response.End
elseif icerik="" and nable=1 or icerik=" " and nable=1 or icerik=Null and nable=1 or icerik="NULL" and nable=1 or icerik="Null" and nable=1 Then
icerik=Null
End If

Set userlogin = Server.CreateObject("ADODB.Recordset")
strSQL ="SELECT * FROM "&table
userlogin.open strSQL,Conne,1,3
userlogin.move(rownum)
userlogin(colnum)=icerik
userlogin.update


if icerik=""or icerik=" " OR icerik="NULL" or isnull(userlogin(colnum))=true Then
Response.Write "NULL"
else
Response.Write trim(userlogin(colnum))
End If
End If

if err.number<>0 Then
response.clear
Response.Write "HATA: "&err.description
End If%>