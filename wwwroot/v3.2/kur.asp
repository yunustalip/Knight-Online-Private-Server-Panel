<!--#include file="ayar.asp"-->
<!--#include file="db.asp"-->

<head>
<meta http-equiv="Content-Language" content="tr">
</head>
<body bgcolor="#F8F8F8">
<%
islem=Request.QueryString("islem")
if islem="yukselt1" then
On Error Resume Next
data.Execute("Create TABLE linkler (id AutoIncrement, isim memo, link memo)")
data.Execute("Create TABLE etiket (id AutoIncrement, blog_id int, etiket memo)")
data.Execute("ALTER TABLE blog ADD yorumdurum int NULL")
data.Execute("UpDate blog Set yorumdurum=0")


data.Execute("ALTER TABLE ayar ADD hakkimda memo NULL")
data.Execute("ALTER TABLE iletisim ADD url memo NULL")

data.Execute("Create TABLE galeri_kat (id AutoIncrement, isim memo, aciklama memo)")
data.Execute("ALTER TABLE galeri ADD kat_id INT NULL")

data.Execute("Create TABLE ankets (id AutoIncrement, soru memo)")
data.Execute("Create TABLE anket (id AutoIncrement, a_id Int, cevap memo, deger Int)")
data.Execute("ALTER TABLE blog ADD gorunsun INT NULL")

Server.ScriptTimeOut=900
Set blgekle = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from galeri_kat"
blgekle.Open SQL,data,1,3
if blgekle.eof then
data.Execute("UpDate galeri Set kat_id=1")

blgekle.Addnew
blgekle("isim")=("Tüm Resimler")
blgekle("aciklama")=("bütün resimler burada toplandi")
blgekle.update
end if

data.Execute("Create TABLE ankets (id AutoIncrement, soru memo)")
data.Execute("Create TABLE anket (id AutoIncrement, a_id Int, cevap memo, deger Int)")
data.Execute("ALTER TABLE blog ADD gorunsun INT NULL")

sql="Select From blog where gorunsun=''"
set blog = data.execute(sql)
if blog.eof then
else
data.Execute("UpDate blog Set gorunsun=1")
end if
blog.close

Function temizle(trh)
	if trh<10 then
		trh=Right(trh,1)
	end if
temizle=trh
End Function

Function Duzelt(duzgun)
if mid(duzgun,3,1)="." then
x=split(duzgun,".")
	ay=x(1)
	gun=x(0)
	yil=x(2)
duzgun= temizle(ay) & "/" & temizle(gun) & "/" & yil
end if
duzelt=duzgun
End Function

Set rs= Server.CreateObject("ADODB.Recordset")
SQL = "Select id,tarih from blog"
rs.Open SQL,data,1,3
do while not rs.eof

rs("tarih") = duzelt(rs("tarih"))
rs.update

rs.movenext
loop
rs.close
set rs=nothing

data.Execute("DROP Table sss") 
data.Close
Set data = Nothing
%>
<div align="center">
	<table border="1" width="500" id="table1" cellspacing="0" cellpadding="0" style="border-collapse: collapse; padding: 10px" bordercolor="#C0C0C0" bgcolor="#FFFFFF">
		<tr>
			<td>
<p align="center">
<font face="Trebuchet MS" style="font-size: 15pt">Ýþlem Tamamlandý</font><br></p>
<font face="Trebuchet MS" style="font-size: 14px">Veritabanýnýz v3.2 ye uygun 
hale getirilmiþtir.<br>
kur.asp yi silmeyi unutmayýnýz<br>
iletiþim: <a href="http://www.webixir.com/iletisim.asp" target="_blank">http://www.webixir.com/iletisim.asp</a></font>
			</td>
		</tr>
	</table>
</div>
<% Else %>
<div align="center">
	<table border="1" width="500" id="table1" cellspacing="0" cellpadding="0" style="border-collapse: collapse; padding: 10px" bordercolor="#C0C0C0" bgcolor="#FFFFFF">
		<tr>
			<td>
			<center><b><font size="5" face="Trebuchet MS">EFENDY BLOG SÜRÜM 
			YÜKSELTME</font></b></center>
			<hr color="#C0C0C0" width="90%" size="1">
			<ul>
				<li><font face="Trebuchet MS" style="font-size: 14px">Ýþlemin 
				amacý veritabanýný v3.2 ye uygun hale getirmektir</font></li>
				<li><font face="Trebuchet MS" style="font-size: 14px">Ýþlem 
				öncesi veritabanýnýzý her ihtimale karþý yedekleyin</font></li>
				<li><font face="Trebuchet MS" style="font-size: 14px">Ýþlem 
				Sonrasý veritabanýnýzýn yapýsý deðiþecektir</font></li>
				<li><font face="Trebuchet MS" style="font-size: 14px">Ýþlem 
				Sonrasýnda veritabanýnýzýn hasarlanmasý veya çökmesi mümkün 
				deðildir. </font> </li>
			</ul>
			<center><font face="Trebuchet MS" style="font-size: 14px">Kurulumu 
			Baþlatmak Ýçin &quot;Kurulumu Baþlat&quot; Butonuna Týklayýnýz.</font></center>
			<br><center>
			<input type="button" value="Kurulumu Baþlat" style="font-size: 14px; font-family: Trebuchet MS; font-weight: bold" onclick="location.href('kur.asp?islem=yukselt1');"></center>
			</td>
		</tr>
	</table>
</div>

<p align="center"><font style="font-size: 11px" face="Trebuchet MS">Bu script Fesih YABAR tarafýndan kodlanmýþ olup 
tamamen ücretsizdir<br>
Sitenin altýndaki <u>EFENDY BLOG </u>yazýsýný silmemeniz rica olunur</font></p>
<% End if %>