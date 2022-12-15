<!--#include file="db.asp"-->
<!--#include file="inc.asp"-->
<!--#include file="filtre.asp"-->
<%
if not isnumeric(Strislemzaman) then Strislemzaman="30"
islem=Request("islem")
if islem="yorumekle" then
id=Request.QueryString("id")

s=Request.QueryString("s")
sid=session.sessionID
if not s = (sid) then
response.redirect("index.asp")
end if

ETarih = Session("SonGon" ) 
If DateDiff("S" ,ETarih,Now()) > Int(Strislemzaman) Then 
Session("SonGon" ) = Now() 

Ekleyen=filtre(request.form("Ekleyen"))
Yorum=filtre(request.form("Yorum"))
	if len(ekleyen) < 3 or len(Yorum) < 3 then
		response.redirect "index.asp"
	end if
Set blgekle = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from yorum"
blgekle.Open SQL,data,1,3

if Stryorumonay="1" then
	tasvip="1"
	cikti="Yorumunuz Admin Tarafýndan Onaylandýktan Sonra Sitede Yerini Alacaktýr."
	if session("admin") then tasvip="0"
else
	tasvip="0"
	cikti="Yorumunuz Baþarýyla Kaydedildi"
end if

blgekle.Addnew
blgekle("Yorum")=Yorum
blgekle("Ekleyen")=Ekleyen
blgekle("Blog_id")=id
blgekle("onay")=tasvip
blgekle("Tarih")=now()
blgekle.update
if session("admin") then
response.redirect Request.ServerVariables("HTTP_REFERER")
end if
Response.Cookies("isim")=Ekleyen
Response.Cookies("isim").Expires = Now() + 7
response.write "<script>alert('"&cikti&"')</script>"
response.write "<meta HTTP-EQUIV=""REFRESH"" content=""0; url="&Request.ServerVariables("HTTP_REFERER")&""">"
response.write "<center><a href="&Request.ServerVariables("HTTP_REFERER")&">Tarayýcýnýz Yönlenmiyorsa Buraya Týklayýn</a></center>"
else
response.write("<script>alert('Ýþlem Zaman Aralýðý "&Strislemzaman&" Sn. den Fazla Olmalý\n\Lutfen "&Strislemzaman&" Sn. Bekleyin.')</script>")
response.write "<meta HTTP-EQUIV=""REFRESH"" content=""0; url=index.asp"">"
response.write "<center><a href=index.asp>Tarayýcýnýz Yönlenmiyorsa Buraya Týklayýn</a></center>"
end if
End if

if islem="oner" then
id=request.querystring("id")
s=request.querystring("s")
sid=session.sessionID
if not s = (sid) then
response.redirect("index.asp")
end if
if Request.Form("isim_git")="" or Request.Form("isim")="" or Request.Form("git")="" then
response.redirect "index.asp"
end if
ETarih = Session("SonGon" ) 
If DateDiff("S" ,ETarih,Now()) > Int(Strislemzaman) Then 
Session("SonGon" ) = Now()
if StrMail="1" then
Set Mail=Server.CreateObject("CDONTS.NewMail")
Mail.To=filtre(Request.Form("git"))
Mail.MailFormat = 1
else
Set Mail=Server.CreateObject("Jmail.Message")
Mail.Charset = "ISO-8859-9"
Mail.AddRecipient filtre(Request.Form("git"))
Mail.contenttype="text/html"
end if
Mail.From=Strtavsiyemail
Mail.Subject="Arkaþýnýzdan Tavsiye - "&Filtre(Request.Form("isim"))
eb=eb &" Selam "& filtre(Request.Form("isim_git")) &","& chr(10)
eb=eb &" "& chr(10)
eb=eb &" Arkadaþýnýz "&Filtre(Request.Form("isim"))&" Size Web Sitemizden Tavsiyede Bulundu"& chr(10)
eb=eb &" Adres: <a href=""http://"&strsite&"/"&SEOLink(id)&""">http://"&strsite&"/"&SEOLink(id)&"</a>" & Chr(10)
eb=eb & chr(10)
eb=eb &" Arkadaþýnýzýn Kýsa Notu: " & chr(10)
eb=eb & filtre(request.form("not")) & chr(10)
eb=Replace(eb,chr(10),"<br />")
if strmail<>1 then
Mail.htmlbody=eb
else
Mail.Body=eb
end if
Mail.Send(strmailserver)
response.write("<script>alert('Tavsiyeniz Arkadaþýnýza Ýletildi');window.close()</script>")
response.write "<center><a href="&Request.ServerVariables("HTTP_REFERER")&">Tarayýcýnýz Yönlenmiyorsa Buraya Týklayýn</a></center>"
else
response.write("<script>alert('Arama Zaman Aralýðý "&Strislemzaman&" Sn. den Fazla Olmalý\n\Lutfen "&Strislemzaman&" Sn. Bekleyin.');window.close()</script>")
end if
end if


if islem="oyver" then

Response.Charset = "iso-8859-9"
Response.Expires=-1

oy = request.querystring("oy")
dokuman = request.querystring("dokuman")

s=Request.QueryString("s")
sid=session.sessionID
if not s = (sid) or isnumeric(oy)=false or isnumeric(dokuman)=false then
response.redirect("index.asp")
end if

if oy >5 then oy=5
if oy <1 then oy=1

set dokumanid = Server.CreateObject("ADODB.Recordset")
sqL = "Select id,deger,degers from blog where id = "&dokuman&""
dokumanid.open sql,data,1,3
if dokumanid.eof or Request.Cookies("Puan")(dokuman) <> "" then
	response.write "bu içeriðe "&Request.Cookies("Puan")(dokuman)&" verilmiþ"
	
else
		dokumanid("deger") 	= dokumanid("deger") + oy
		dokumanid("degers")	= dokumanid("degers") + 1
		dokumanid.update
			Session("puan") = "1"
		Response.Cookies("Puan")(dokuman) = oy
		Response.Cookies("Puan").Expires = now()+1
		response.write "<font class=orta>"&oy&" verildi</font>"
end if
end if
%>
<%
if islem="yazdir" then
id=request.querystring("id")
	if id="" then
		response.redirect "index.asp"
	end if
s=Request.QueryString("s")
sid=session.sessionID
if not s = (sid) then
response.redirect("index.asp")
end if

set yaz = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from blog where id like '"&id&"'"
yaz.open SQL,data,1,3
if yaz.eof or yaz.bof then
response.redirect "index.asp"
end if
isim = yaz("konu")
icerik = yaz("mesaj")
tarih = yaz("tarih")
%>
<%
    Response.ContentType = "application/msword"
    Response.AddHeader "Content-Disposition", "attachment;filename="&isim&".doc"
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=Windows-1254">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 11">
<meta name=Originator content="Microsoft Word 11">
<title>Blog</title>
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
  <o:Author>Efendy</o:Author>
  <o:Template>Normal</o:Template>
  <o:LastAuthor>Efendy</o:LastAuthor>
  <o:Revision>1</o:Revision>
  <o:TotalTime>2</o:TotalTime>
  <o:Created>2007-06-20T12:07:00Z</o:Created>
  <o:LastSaved>2007-06-20T12:09:00Z</o:LastSaved>
  <o:Pages>1</o:Pages>
  <o:Words>2</o:Words>
  <o:Characters>18</o:Characters>
  <o:Lines>1</o:Lines>
  <o:Paragraphs>1</o:Paragraphs>
  <o:CharactersWithSpaces>19</o:CharactersWithSpaces>
  <o:Version>11.5606</o:Version>
 </o:DocumentProperties>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <w:WordDocument>
  <w:SpellingState>Clean</w:SpellingState>
  <w:GrammarState>Clean</w:GrammarState>
  <w:HyphenationZone>21</w:HyphenationZone>
  <w:PunctuationKerning/>
  <w:ValidateAgainstSchemas/>
  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>
  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
  <w:Compatibility>
   <w:BreakWrappedTables/>
   <w:SnapToGridInCell/>
   <w:WrapTextWithPunct/>
   <w:UseAsianBreakRules/>
   <w:DontGrowAutofit/>
  </w:Compatibility>
  <w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel>
 </w:WordDocument>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <w:LatentStyles DefLockedState="false" LatentStyleCount="156">
 </w:LatentStyles>
</xml><![endif]-->
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:Tahoma;
	panose-1:2 11 6 4 3 5 4 4 2 4;
	mso-font-charset:162;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:1627421319 -2147483648 8 0 66047 0;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
h1
	{mso-style-next:Normal;
	margin-top:12.0pt;
	margin-right:0cm;
	margin-bottom:3.0pt;
	margin-left:0cm;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:1;
	font-size:16.0pt;
	font-family:Arial;
	mso-font-kerning:16.0pt;}
@page Section1
	{size:595.3pt 841.9pt;
	margin:70.85pt 70.85pt 70.85pt 70.85pt;
	mso-header-margin:35.4pt;
	mso-footer-margin:35.4pt;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
-->
</style>
<!--[if gte mso 10]>
<style>
 /* Style Definitions */
 table.MsoNormalTable
	{mso-style-name:"Normal Tablo";
	mso-tstyle-rowband-size:0;
	mso-tstyle-colband-size:0;
	mso-style-noshow:yes;
	mso-style-parent:"";
	mso-padding-alt:0cm 5.4pt 0cm 5.4pt;
	mso-para-margin:0cm;
	mso-para-margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	font-family:"Times New Roman";
	mso-ansi-language:#0400;
	mso-fareast-language:#0400;
	mso-bidi-language:#0400;}
</style>
<![endif]-->
</head>

<body lang=TR style='tab-interval:35.4pt'>

<div class=Section1>

<h1><%=yaz("konu")%></h1><br>
<span style='font-size:11px;font-family:Tahoma'><%=Replace(yaz("mesaj"),"{KES}","")%><br>
<br>
<span style='color:red'>Eklenme: <%=yaz("tarih")%></span></span>

</div>

</body>

</html>

<%
yaz.Close
Set yaz = Nothing
End if

if islem="anket" then

s=Request.QueryString("s")
sid=session.sessionID
if not s = (sid) then
response.redirect("index.asp")
end if

pid = Request.QueryString("id")
cevap = Request.Form("cevap")
if cevap="" then
response.redirect "index.asp"
end if
gelen = Request.ServerVariables("HTTP_REFERER")
if Request.Cookies("anket")(pid) = "kapat" then
response.write "<script>alert('Önceden Oy Kullandýnýz.');</script>"
response.write "<meta HTTP-EQUIV=""REFRESH"" content=""0; url="&Request.ServerVariables("HTTP_REFERER")&""">"
response.write "<center><a href="&Request.ServerVariables("HTTP_REFERER")&">Tarayýcýnýz Yönlenmiyorsa Buraya Týklayýn</a></center>"
else
SQL ="SELECT * FROM anket WHERE id=" & cevap
set anket = Server.CreateObject("ADODB.RecordSet")
anket.Open SQL, data, 1, 3
if anket.eof or anket.bof then
response.redirect "index.asp"
end if
anket("deger")=int(anket("deger")) + 1
anket.update
   Response.Cookies("anket")(pid) = "kapat" 'Bu kýsým güvenlik önlemi. Adamýn bilgisayarýna cookie atýyoruz
   Response.Cookies("anket").Expires = Now() + 365 '365 gün boyunca ayný ankete oy kullanmamasý için :)
response.write "<script>alert('Oy Kullandiginiz Ýçin Teþekkür Ederiz');</script>"
response.write "<meta HTTP-EQUIV=""REFRESH"" content=""0; url="&Request.ServerVariables("HTTP_REFERER")&""">"
response.write "<center><a href="&Request.ServerVariables("HTTP_REFERER")&">Tarayýcýnýz Yönlenmiyorsa Buraya Týklayýn</a></center>"
end if
end if
%>
<%
if islem="tavsiyeet" then
id=Request.QueryString("id")
	if isnumeric(id) then
Set rs = Server.CreateObject("Adodb.Recordset")
SQL="Select id from blog where id="&id&""
rs.open SQL,data,1,3
if not rs.eof then
%>
<link href="tema/stil.css" rel="stylesheet" type="text/css">
<style>
body {background-color:#fffff}
</style>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
<SCRIPT LANGUAGE=JAVASCRIPT>
function validate(form) {
if (form.isim.value == "") {
   alert("Adýnýzý Yazýnýz.");
   return false; }
if (form.isim_git.value == "") {
   alert("Arkadaþýnýzýn Adýný Yazýnýz");
   return false; }
if (form.git.value == "") {
   alert("Arkadaþýnýzýn Mail Adresini Yazýnýz");
   return false; }
return true;
}
</SCRIPT>
<table border="0" width="100%" height="100%" cellpadding="0" class="js_tool">
<form action="islem.asp?islem=oner&id=<%=id%>&s=<%=session.sessionID%>" method="post" onSubmit="return validate(this)">
	<tr>
		<td width="30%" align="right"><font class="orta">Adýn :</font></td>
		<td width="70%"> 
        <input type=text name="isim" style="width:100%" class="alan"></td>
	</tr>
	<tr>
		<td width="100%" colspan="2" class="tool">
		<p align="center"><b><font class="orta">Arkadaþýnýn</font></b></td>
		</tr>
	<tr>
		<td width="30%" align="right"><font class="orta">Adý :</font></td>
		<td width="70%"> 
          <input type=text name="isim_git" style="width:100%" class="alan"></td>
	</tr>
	<tr>
		<td width="30%" align="right"><font class="orta">Mail Adresi :</font></td>
		<td width="70%"> 
          <input type="text" name="git" style="width:100%" class="alan"></td>
	</tr>
	<tr>
		<td colspan="2"><font class="orta">Kýsa Not:<br></font>
		<textarea rows="4" class="alan" name="not" onKeyUp="return ismaxlength(this)" maxlength="300" style="width:100%"></textarea>
		</td>
	</tr>
	<tr>
		<td width="100%" colspan="2">
         <input type="submit" value="Gönder" class="dugme"></td>
	</tr>
</form>
</table>
<%
end if 
rs.close : set rs=nothing
end if
end if

if islem="ilet" then
if not isnumeric(Strislemzaman) then Strislemzaman="30"
s=Request.QueryString("s")
sid=session.sessionID
if not s = (sid) then
response.redirect("index.asp")
end if

ETarih = Session("SonGon" ) 
If DateDiff("S" ,ETarih,Now()) > Int(Strislemzaman) Then 
Session("SonGon" ) = Now()

Konu=filtre(request.form("Konu"))
Mesaj=filtre(request.form("Mesaj"))
isim=filtre(request.form("isim"))
yer=filtre(request.form("yer"))
mail=filtre(request.form("mail"))
url=filtre(request.form("url"))
	if len(konu) < 3 or Len(Mesaj) < 3 or Len(isim) < 3 or Len(yer) < 3 or Len(mail) < 7 then
		response.redirect "index.asp"
	end if
Set blgekle = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from iletisim"
blgekle.Open SQL,data,1,3

blgekle.Addnew
blgekle("Mesaj")=Mesaj
blgekle("Konu")=Konu
blgekle("isim")=isim
blgekle("yer")=yer
blgekle("mail")=mail
blgekle("url")=url
blgekle("Tarih")=now()
blgekle.update
Response.Cookies("isim")=isim
Response.Cookies("isim").Expires = Now() + 7
Response.Cookies("yer")=yer
Response.Cookies("yer").Expires = Now() + 7
Response.Cookies("mail")=mail
Response.Cookies("mail").Expires = Now() + 7
response.write("<script>alert('Yorumunuz Site Yöneticisine Ulaþmýþtýr')</script>")
response.write("<script>location.href('index.asp')</script>")
else
response.write("<script>alert('Ýþlem Zaman Aralýðý "&Strislemzaman&" Sn. den Fazla Olmalý\n\Lutfen "&Strislemzaman&" Sn. Bekleyin.')</script>")
response.write("<script>location.href('index.asp')</script>")
end if
end if

if islem="yaz" then
if not isnumeric(Strislemzaman) then Strislemzaman="30"
s=Request.QueryString("s")
sid=session.sessionID
if not s = (sid) then
response.redirect("index.asp")
end if
ETarih = Session("SonGon" ) 
If DateDiff("S" ,ETarih,Now()) > Int(Strislemzaman) Then 
Session("SonGon" ) = Now()
mail=filtre(request.form("mail"))
Mesaj=filtre(request.form("Mesaj"))
yazan=filtre(request.form("yazan"))
yer=filtre(request.form("yer"))

aid=Filtre(Request.QueryString("aid"))
if not aid="" then
if isnumeric(aid)=false then
response.redirect"index.asp"
else
set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from zd where id="&Cint(Filtre(aid))&""
rs.open SQL,data,1,3
if not rs.eof then
mesaj="<div class=alinti><b>Alýntý Sahibi: "&rs("yazan")&"</b>"&vBCrlf&rs("mesaj")&"</div>"&vBCrlf&mesaj
end if
rs.close : set rs=nothing
end if
end if
	if len(mail) < 3 or Len(Mesaj) < 3 or Len(yazan) < 3 or Len(yer) < 3 then
		response.redirect "index.asp"
	end if
if Strzdonay="1" then
	tasvip="1"
	cikti="Mesajýnýz Admin Tarafýndan Onaylandýktan Sonra Sitede Yerini Alacaktýr."
	if session("admin") then tasvip="0"
else
	tasvip="0"
	cikti="Mesajýnýz Baþarýyla Kaydedildi"
end if

Set blgekle = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from zd"
blgekle.Open SQL,data,1,3

blgekle.Addnew
blgekle("Mesaj")=Mesaj
blgekle("mail")=mail
blgekle("yazan")=yazan
blgekle("yer")=yer
blgekle("onay")=tasvip
blgekle("Tarih")=now()
blgekle.update
blgekle.close
set blgekle = nothing

Response.Cookies("isim")=yazan
Response.Cookies("isim").Expires = Now() + 7
Response.Cookies("yer")=yer
Response.Cookies("yer").Expires = Now() + 7
Response.Cookies("mail")=mail
Response.Cookies("mail").Expires = Now() + 7

if session("admin") then
response.redirect("zd.asp")
end if

response.write("<script>alert('"&cikti&"')</script>")
response.write("<script>location.href('zd.asp')</script>")
else
response.write("<script>alert('Ýþlem Zaman Aralýðý "&Strislemzaman&" Sn. den Fazla Olmalý\n\Lutfen "&Strislemzaman&" Sn. Bekleyin.')</script>")
response.write("<script>location.href('index.asp')</script>")
end if
end if
%>