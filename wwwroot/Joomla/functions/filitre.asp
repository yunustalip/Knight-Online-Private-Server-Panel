<%
'      JoomlASP Site Yönetimi Sistemi (CMS)
'
'      Copyright (C) 2007 Hasan Emre ASKER
'
'      This program is free software; you can redistribute it and/or modify it
'      under the terms of the GNU General Public License as published by the Free
'      Software Foundation; either version 3 of the License, or (at your option)
'      any later version.
'
'      This program is distributed in the hope that it will be useful, but WITHOUT
'      ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
'      FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'
'      You should have received a copy of the GNU General Public License along with
'      this library; if not, write to the JoomlASP Asp Yazýlým Sistemleri., Kargaz Doðal Gaz Bilgi Ýþlem Müdürlüðü
'       36100 Kars / Merkez 
'		Tel : 0544 275 9804 - 0537 275 3655
'		Mail: emre06@hotmail.com.tr / info@joomlasp.com/.net/.org
'
'
'		Lisans Anlaþmasý Gereði Lütfen Google Reklam Bölümünü Sitenizden kaldýrmayýnýz. Bu sizin GOOGLE reklamlarýný yapmanýza
'		kesinlikle bir engel deðildir. reklam.asp içeriðinin yada yayýnladýðý verinin deðiþmesi lisans politikasýnýn dýþýna çýkýlmasýna
'		ve JoomlASP CMS sistemini ücretsiz yayýnlamak yerine ücretlie hale getirmeye bizi teþfik etmektedir. Bu Sistem için verilen emeðe
'		saygý ve bir çeþit ödeme seçeneði olarak GOOGLE reklamýmýzýn deðiþtirmemesi yada silinmemesi gerekmektedir.
%>
<%
function guvenlik(data) 
'Güvenlik Bölgesi
data = Replace(data, "%22", "", 1, -1, 1)
data = Replace(data, "%27", "", 1, -1, 1)
data = Replace(data, "%0", "", 1, -1, 1)
data = Replace(data, "\", "\\", 1, -1, 1)
data = Replace(data, "'", "\'", 1, -1, 1)
data = Replace(data, "[", "&#091;", 1, -1, 1)
data = Replace(data, "]", "&#093;", 1, -1, 1)
data = Replace(data, "<", "&lt;", 1, -1, 1)
data = Replace(data, ">", "&gt;", 1, -1, 1)
data = Replace (data ,"`","",1,-1,1) 
data = Replace (data ,"=","",1,-1,1) 
data = Replace (data ,"&","",1,-1,1) 
data = Replace (data ,"%","",1,-1,1) 
data = Replace (data ,"!","",1,-1,1) 
data = Replace (data ,"#","",1,-1,1) 
data = Replace (data ,"<","",1,-1,1) 
data = Replace (data ,">","",1,-1,1) 
data = Replace (data ,"*","",1,-1,1) 
data = Replace (data ,"And","",1,-1,1) 
data = Replace (data ,"select","",1,-1,1) 
data = Replace (data ,"join","",1,-1,1) 
data = Replace (data ,"union","",1,-1,1) 
data = Replace (data ,"where","",1,-1,1)
data = Replace (data ,"execute","",1,-1,1) 
data = Replace (data ,"insert","",1,-1,1) 
data = Replace (data ,"delete","",1,-1,1) 
data = Replace (data ,"update","",1,-1,1) 
data = Replace (data ,"like","",1,-1,1) 
data = Replace (data ,"drop","",1,-1,1) 
data = Replace (data ,"create","",1,-1,1) 
data = Replace (data ,"modify","",1,-1,1) 
data = Replace (data ,"rename","",1,-1,1) 
data = Replace (data ,"alter","",1,-1,1) 
data = Replace (data ,"cast","",1,-1,1)

guvenlik=data 
end function

function guvenlikyorum(veri)
	veri = Replace(veri, "[", "&#091;", 1, -1, 1)
	veri = Replace(veri, "]", "&#093;", 1, -1, 1)
	veri = Replace(veri, "<", "&lt;", 1, -1, 1)
	veri = Replace(veri, ">", "&gt;", 1, -1, 1)
	veri = Replace(veri, "<script language=""javascript"">", "", 1, -1, 1)
	veri = Replace(veri, "<script language=javascript>", "", 1, -1, 1)
	veri = Replace(veri, "<script language=""vbscript"">", "", 1, -1, 1)
	veri = Replace(veri, "<script language=vbscript>", "", 1, -1, 1)
	veri = Replace(veri, "<script language=""jscript"">", "", 1, -1, 1)
	veri = Replace(veri, "<script language=jscript>", "", 1, -1, 1)
	veri = Replace(veri, "<script type=""text/javascript"">", "", 1, -1, 1)
	veri = Replace(veri, "<script type=text/javascript>", "", 1, -1, 1)
	veri = Replace(veri, "<script type=""text/vbscript"">", "", 1, -1, 1)
	veri = Replace(veri, "<script type=text/vbscript>", "", 1, -1, 1)
	veri = Replace(veri, "<script type=""text/jscript"">", "", 1, -1, 1)
	veri = Replace(veri, "<script type=text/jscript>", "", 1, -1, 1)
	veri = Replace(veri, "<script>", "", 1, -1, 1)
	veri = Replace(veri, "</script>", "", 1, -1, 1)
	veri = Replace(veri, "<style", "<", 1, -1, 1)
	veri = Replace(veri, "</style>", "", 1, -1, 1)
	veri = Replace(veri ,"'","''",1,-1,1)
	veri = Replace(veri, "%22", "", 1, -1, 1)
	veri = Replace(veri, "%27", "", 1, -1, 1)
	guvenlikyorum=veri
end function

Function uyeisimkontrol(kontroller)

	kontroller = Replace(kontroller, "'", "", 1, -1, 1)
	kontroller = Replace(kontroller, "[", "", 1, -1, 1)
	kontroller = Replace(kontroller, "]", "", 1, -1, 1)
	kontroller = Replace(kontroller, "<", "", 1, -1, 1)
	kontroller = Replace(kontroller, ">", "", 1, -1, 1)
	kontroller = Replace(kontroller ,"`","",1,-1,1) 
	kontroller = Replace(kontroller ,"=","",1,-1,1) 
	kontroller = Replace(kontroller ,"&","",1,-1,1) 
	kontroller = Replace(kontroller ,"%","",1,-1,1) 
	kontroller = Replace(kontroller ,"!","",1,-1,1) 
	kontroller = Replace(kontroller ,"#","",1,-1,1) 
	kontroller = Replace(kontroller ,"<","",1,-1,1) 
	kontroller = Replace(kontroller ,">","",1,-1,1) 
	kontroller = Replace(kontroller ,"*","",1,-1,1)
	kontroller = Replace(kontroller, Chr(9), "", 1, -1, 1)
	kontroller = Replace(kontroller, "</script>", "", 1, -1, 1)
	kontroller = Replace(kontroller, "<script language=""javascript"">", "", 1, -1, 1)
	kontroller = Replace(kontroller, "<script language=javascript>", "", 1, -1, 1)
	kontroller = Replace(kontroller, "SCRIPT", "&#083;CRIPT", 1, -1, 0)
	kontroller = Replace(kontroller, "Script", "&#083;cript", 1, -1, 0)
	kontroller = Replace(kontroller, "script", "&#115;cript", 1, -1, 1)
	kontroller = Replace(kontroller, "MOCHA", "&#077;OCHA", 1, -1, 0)
	kontroller = Replace(kontroller, "Mocha", "&#077;ocha", 1, -1, 0)
	kontroller = Replace(kontroller, "mocha", "&#109;ocha", 1, -1, 1)
	kontroller = Replace(kontroller, "OBJECT", "&#079;BJECT", 1, -1, 0)
	kontroller = Replace(kontroller, "Object", "&#079;bject", 1, -1, 0)
	kontroller = Replace(kontroller, "object", "&#111;bject", 1, -1, 1)
	kontroller = Replace(kontroller, "APPLET", "&#065;PPLET", 1, -1, 0)
	kontroller = Replace(kontroller, "Applet", "&#065;pplet", 1, -1, 0)
	kontroller = Replace(kontroller, "applet", "&#097;pplet", 1, -1, 1)
	kontroller = Replace(kontroller, "ALERT", "&#065;LERT", 1, -1, 0)
	kontroller = Replace(kontroller, "Alert", "&#065;lert", 1, -1, 0)
	kontroller = Replace(kontroller, "alert", "&#097;lert", 1, -1, 1)
	kontroller = Replace(kontroller, "EMBED", "&#069;MBED", 1, -1, 0)
	kontroller = Replace(kontroller, "Embed", "&#069;mbed", 1, -1, 0)
	kontroller = Replace(kontroller, "embed", "&#101;mbed", 1, -1, 1)
	kontroller = Replace(kontroller, "EVENT", "&#069;VENT", 1, -1, 0)
	kontroller = Replace(kontroller, "Event", "&#069;vent", 1, -1, 0)
	kontroller = Replace(kontroller, "event", "&#101;vent", 1, -1, 1)
	kontroller = Replace(kontroller, "DOCUMENT", "&#068;OCUMENT", 1, -1, 0)
	kontroller = Replace(kontroller, "Document", "&#068;ocument", 1, -1, 0)
	kontroller = Replace(kontroller, "document", "&#100;ocument", 1, -1, 1)
	kontroller = Replace(kontroller, "COOKIE", "&#067;OOKIE", 1, -1, 0)
	kontroller = Replace(kontroller, "Cookie", "&#067;ookie", 1, -1, 0)
	kontroller = Replace(kontroller, "cookie", "&#099;ookie", 1, -1, 1)
	kontroller = Replace(kontroller, "IFRAME", "I&#070;RAME", 1, -1, 0)
	kontroller = Replace(kontroller, "Iframe", "I&#102;rame", 1, -1, 0)
	kontroller = Replace(kontroller, "iframe", "i&#102;rame", 1, -1, 1)
	kontroller = Replace(kontroller, "TEXTAREA", "&#84;EXTAREA", 1, -1, 0)
	kontroller = Replace(kontroller, "Textarea", "&#84;extarea", 1, -1, 0)
	kontroller = Replace(kontroller, "textarea", "&#116;extarea", 1, -1, 1)
	kontroller = Replace(kontroller, "<STR&#079;NG>", "<strong>", 1, -1, 1)
	kontroller = Replace(kontroller, "<str&#111;ng>", "<strong>", 1, -1, 1)
	kontroller = Replace(kontroller, "</STR&#079;NG>", "</strong>", 1, -1, 1)
	kontroller = Replace(kontroller, "</str&#111;ng>", "</strong>", 1, -1, 1)
	kontroller = Replace(kontroller, "f&#111;nt", "font", 1, -1, 0)
	kontroller = Replace(kontroller, "F&#079;NT", "FONT", 1, -1, 0)
	kontroller = Replace(kontroller, "F&#111;nt", "Font", 1, -1, 0)
	kontroller = Replace(kontroller, "f&#079;nt", "font", 1, -1, 1)
	kontroller = Replace(kontroller, "f&#111;nt", "font", 1, -1, 1)
	kontroller = Replace(kontroller, "m&#111;no", "mono", 1, -1, 0)
	kontroller = Replace(kontroller, "M&#079;NO", "MONO", 1, -1, 0)
	kontroller = Replace(kontroller, "M&#111;no", "Mono", 1, -1, 0)
	kontroller = Replace(kontroller, "m&#079;no", "mono", 1, -1, 1)
	kontroller = Replace(kontroller, "m&#111;no", "mono", 1, -1, 1)
	uyeisimkontrol = kontroller
End Function


function guvenmesajoku(mesajoku) 
'Güvenlik Bölgesi


mesajoku = Replace (mesajoku, "<", "&lt;", 1, -1, 1)
mesajoku = Replace (mesajoku, ">", "&gt;", 1, -1, 1)
mesajoku = Replace (mesajoku, "[b]", "<b>", 1, -1, 1)
mesajoku = Replace (mesajoku, "[/b]", "</b>", 1, -1, 1)
mesajoku = Replace (mesajoku, "[img]", "<img width=""300"" height=""350"" src=", 1, -1, 1)
mesajoku = Replace (mesajoku, "[/img]", " />", 1, -1, 1)
mesajoku = Replace (mesajoku, Chr(13), "<br>", 1, -1, 1)
mesajoku = Replace (mesajoku, "<script language=""javascript"">", "", 1, -1, 1)
mesajoku = Replace (mesajoku, "<script language=javascript>", "", 1, -1, 1)
mesajoku = Replace (mesajoku, "<script language=""vbscript"">", "", 1, -1, 1)
mesajoku = Replace (mesajoku, "<script language=vbscript>", "", 1, -1, 1)
mesajoku = Replace (mesajoku, "<script language=""jscript"">", "", 1, -1, 1)
mesajoku = Replace (mesajoku, "<script language=jscript>", "", 1, -1, 1)
mesajoku = Replace (mesajoku, "<script type=""text/javascript"">", "", 1, -1, 1)
mesajoku = Replace (mesajoku, "<script type=text/javascript>", "", 1, -1, 1)
mesajoku = Replace (mesajoku, "<script type=""text/vbscript"">", "", 1, -1, 1)
mesajoku = Replace (mesajoku, "<script type=text/vbscript>", "", 1, -1, 1)
mesajoku = Replace (mesajoku, "<script type=""text/jscript"">", "", 1, -1, 1)
mesajoku = Replace (mesajoku, "<script type=text/jscript>", "", 1, -1, 1)

guvenmesajoku=mesajoku 
end function

function guvenmesajyaz(mesajyaz) 
'Güvenlik Bölgesi
mesajyaz = Replace (mesajyaz, "'", "", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<br>", Chr(13), 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<img width=""300"" height=""350"" src=", "[img]", 1, -1, 1)
mesajyaz = Replace (mesajyaz, " />", "[/img]", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<", "&lt;", 1, -1, 1)
mesajyaz = Replace (mesajyaz, ">", "&gt;", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<b>", "[b]", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "</b>", "[/b]", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<script language=""javascript"">", "", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<script language=javascript>", "", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<script language=""vbscript"">", "", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<script language=vbscript>", "", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<script language=""jscript"">", "", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<script language=jscript>", "", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<script type=""text/javascript"">", "", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<script type=text/javascript>", "", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<script type=""text/vbscript"">", "", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<script type=text/vbscript>", "", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<script type=""text/jscript"">", "", 1, -1, 1)
mesajyaz = Replace (mesajyaz, "<script type=text/jscript>", "", 1, -1, 1)

guvenmesajyaz=mesajyaz 
end function

function dilkontrol(dilim) 
'Güvenlik Bölgesi
dilim = Replace (dilim, "%22", "", 1, -1, 1)
dilim = Replace (dilim, "%27", "", 1, -1, 1)
dilim = Replace (dilim, "%0", "", 1, -1, 1)
dilim = Replace (dilim, "\", "\\", 1, -1, 1)
dilim = Replace (dilim, "'", "'", 1, -1, 1)
dilim = Replace (dilim, "[", "&#091;", 1, -1, 1)
dilim = Replace (dilim, "]", "&#093;", 1, -1, 1)
dilim = Replace (dilim, "<", "&lt;", 1, -1, 1)
dilim = Replace (dilim, ">", "&gt;", 1, -1, 1)
dilim = Replace (dilim, "`", "", 1,-1,1) 
dilim = Replace (dilim, "=", "=", 1,-1,1) 
dilim = Replace (dilim, "&", "", 1,-1,1) 
dilim = Replace (dilim, "%", "", 1,-1,1) 
dilim = Replace (dilim, "!", "", 1,-1,1) 
dilim = Replace (dilim, "#", "'", 1,-1,1) 
dilim = Replace (dilim, "*", "", 1,-1,1) 
dilkontrol=dilim 
end function

Function Virgulle(rakamlar)
    rakamlar = Cstr(StrReverse(rakamlar))
For i = 1 To Len(rakamlar)
    GeciciVeri = Mid(rakamlar,i,1) & GeciciVeri
If i Mod 3 = 0 And Not i = Len(rakamlar) Then GeciciVeri = "," & GeciciVeri
Next
    Virgulle = GeciciVeri
End Function


Function SifreUret(Uzunluk)
Karakterler = "0123456789"
Randomize
KarakterBoyu = Len(Karakterler)
For i = 1 To Uzunluk
      KacinciKarakter = Int((KarakterBoyu * Rnd) + 1)
      UretilenSifre = UretilenSifre & Mid(Karakterler,KacinciKarakter,1)
Next
SifreUret = UretilenSifre
End Function
secure_code = SifreUret(4)
%>
