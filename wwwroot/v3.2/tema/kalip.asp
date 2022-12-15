<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title><%=baslik%> - <%=sitebaslik%></title>
<meta name="description" content="<%=aciklama%>">
<meta name="keywords" content="<%=etiket%>">
<link rel="shortcut icon" href="favicon.ico">
<link rel="alternate" type="application/rss+xml" title="Son Yorumlar - <%=sitebaslik%>" href="http://<%=strsite%>/Rss.asp?rss=yorumlar"/>
<link rel="alternate" type="application/rss+xml" title="Son Bloglar - <%=sitebaslik%>" href="http://<%=strsite%>/Rss.asp?rss=bloglar"/>
<link href="tema/stil.css" rel="stylesheet" type="text/css"><% if Instr(1,script,"404.asp",1) or Instr(1,script,"blog.asp",1)  or Instr(1,script,"hakkimda.asp",1) then %>
<script type="text/javascript" src="ajax/ajax_navigation.js"></script>
<link href="ajax/alt_star_rating.css" rel="stylesheet" type="text/css" media="all">
<script src="ajax/xmlhttp.js" type="text/javascript"></script>
<script src="ajax/rating.js" type="text/javascript"></script><% End if %>
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0"<% if Instr(1,script,"404.asp",1) or Instr(1,script,"blog.asp",1) or Instr(1,script,"hakkimda.asp",1) then %> onload="open_url('_yorum.asp?id=<%=id1%>','my_site_content');"<%end if%>>
<div align="center">
	<table border="0" width="830" id="table1" cellpadding="0" style="border-collapse: collapse">
		<tr>
			<td width="34">
			<img border="0" src="tema/v3_img/v3_02.gif" width="127" height="95"></td>
			<td width="411" background="tema/v3_img/v3_03.gif" align="center">

<!--span-->
<span id="AS" style="display: none;"><font class="menu">ANASAYFA / BLOG</font></span>
<span id="RG" style="display: none;"><font class="menu">RESİM GALERİSİ</font></span>
<span id="ZD" style="display: none;"><font class="menu">ZİYARETÇİ DEFTERİ</font></span>
<span id="IL" style="display: none;"><font class="menu">İLETİŞİM</font></span>
<span id="HK" style="display: none;"><font class="menu">HAKKIMDA</font></span>
<!--/span-->
			
			</td>
			<td width="60">
			<a href="index.asp" onmouseout="javascript:document.getElementById('AS').style.display = (document.getElementById('AS').style.display == 'none' ? 'block' : 'none');" onmouseover="javascript:document.getElementById('AS').style.display = (document.getElementById('AS').style.display == 'none' ? 'block' : 'none');">
			<img border="0" src="tema/v3_img/v3_04.gif" width="77" height="95"></a></td>
			<td width="55">
			<a href="galeri.asp" onmouseout="javascript:document.getElementById('RG').style.display = (document.getElementById('RG').style.display == 'none' ? 'block' : 'none');" onmouseover="javascript:document.getElementById('RG').style.display = (document.getElementById('RG').style.display == 'none' ? 'block' : 'none');">
			<img border="0" src="tema/v3_img/v3_05.gif" width="76" height="95"></a></td>
			<td width="50">
			<a href="zd.asp" onmouseout="javascript:document.getElementById('ZD').style.display = (document.getElementById('ZD').style.display == 'none' ? 'block' : 'none');" onmouseover="javascript:document.getElementById('ZD').style.display = (document.getElementById('ZD').style.display == 'none' ? 'block' : 'none');">
			<img border="0" src="tema/v3_img/v3_06.gif" width="81" height="95"></a></td>
			<td width="62">
			<a href="iletisim.asp" onmouseout="javascript:document.getElementById('IL').style.display = (document.getElementById('IL').style.display == 'none' ? 'block' : 'none');" onmouseover="javascript:document.getElementById('IL').style.display = (document.getElementById('IL').style.display == 'none' ? 'block' : 'none');">
			<img border="0" src="tema/v3_img/v3_07.gif" width="75" height="95"></a></td>
			<td width="42">
			<a href="hakkimda.asp" onmouseout="javascript:document.getElementById('HK').style.display = (document.getElementById('HK').style.display == 'none' ? 'block' : 'none');" onmouseover="javascript:document.getElementById('HK').style.display = (document.getElementById('HK').style.display == 'none' ? 'block' : 'none');">
			<img border="0" src="tema/v3_img/v3_08.gif" width="76" height="95"></a></td>
			<td width="23">
			<img border="0" src="tema/v3_img/v3_09.gif" width="27" height="95"></td>
		</tr>
	</table>
</div>
<div align="center">
	<table border="0" width="830" id="table2" cellpadding="0" style="border-collapse: collapse">
		<tr>
			<td width="17">
			<img border="0" src="tema/v3_img/v3_11.gif" width="28" height="28"></td>
			<td width="16">
			<a onmouseout="javascript:document.getElementById('ay').style.display = (document.getElementById('ay').style.display == 'none' ? 'block' : 'none');" onmouseover="javascript:document.getElementById('ay').style.display = (document.getElementById('ay').style.display == 'none' ? 'block' : 'none');" style="CURSOR:hand;" href="javascript:void(0)" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('http://<%=strsite%>');"><img border="0" src="tema/v3_img/son-hali_05.gif" width="22" height="28"></a></td>
			<td width="16">
			<a onmouseout="javascript:document.getElementById('fy').style.display = (document.getElementById('fy').style.display == 'none' ? 'block' : 'none');" onmouseover="javascript:document.getElementById('fy').style.display = (document.getElementById('fy').style.display == 'none' ? 'block' : 'none');" href="javascript:window.external.AddFavorite('http://<%=strsite%>','<%=sitebaslik%>');"><img border="0" src="tema/v3_img/v3_13.gif" width="26" height="28"></a></td>
			<td width="15">
			<a onmouseout="javascript:document.getElementById('pl').style.display = (document.getElementById('pl').style.display == 'none' ? 'block' : 'none');" onmouseover="javascript:document.getElementById('pl').style.display = (document.getElementById('pl').style.display == 'none' ? 'block' : 'none');" href="paylas.asp?baslik=<%=baslik%>"><img border="0" src="tema/v3_img/v3_14.gif" width="25" height="28" border="0"></a></td>
			<td width="12">
			<a onmouseout="javascript:document.getElementById('rs').style.display = (document.getElementById('rs').style.display == 'none' ? 'block' : 'none');" onmouseover="javascript:document.getElementById('rs').style.display = (document.getElementById('rs').style.display == 'none' ? 'block' : 'none');" target="_blank" href="rss.asp"><img border="0" src="tema/v3_img/v3_15.gif" width="26" height="28"></a></td>
			<td width="24">
			<a onmouseout="javascript:document.getElementById('ar').style.display = (document.getElementById('ar').style.display == 'none' ? 'block' : 'none');" onmouseover="javascript:document.getElementById('ar').style.display = (document.getElementById('ar').style.display == 'none' ? 'block' : 'none');" href="arsiv.asp" target="_blank"><img border="0" src="tema/v3_img/arsiv.gif" width="24" height="28"></a></td>
			<td width="455" background="tema/v3_img/v3_16.gif">
<!--span2-->
<span id="ay" style="display: none;"><font class="menuk"> Anasayfam Yap</font></span>
<span id="fy" style="display: none;"><font class="menuk"> Favorim Yap</font></span>
<span id="pl" style="display: none;"><font class="menuk"> Paylaş</font></span>
<span id="rs" style="display: none;"><font class="menuk"> RSS</font></span>
<span id="ar" style="display: none;"><font class="menuk"> Blog Arşivi</font></span>
<!--/span2-->
			</td>
<form action="ara.asp" method="get">
			<td width="221" background="tema/v3_img/v3_16.gif" align="right">
<input type="text" value="<%=kelime%>" size="23" style="font-size: 9pt; font-family: Tahoma; font-weight:bold; color: #FFFFFF; border: 1px solid gray; background:transparent;" name="ara">&nbsp;<input type="submit" value="Ara" style="font-size: 9pt; font-family: Tahoma; font-weight:bold; color: #FFFFFF; border: 1px solid gray; background:transparent;">
			</td>
</form>
			<td width="16">
			<img border="0" src="tema/v3_img/v3_17.gif" width="27" height="28"></td>
		</tr>
		<tr>
			<td colspan="9" height="13" background="tema/v3_img/son-hali_03.gif"></td>
		</tr>
	</table>
</div>

<div align="center">
	<table border="0" width="830" id="table3" cellpadding="0" style="border-collapse: collapse">
		<tr>
			<td width="830" colspan="3" bgcolor="#ffffff" background="tema/v3_img/govde_orta.gif">

<div align="center">
	<table border="0" width="770" id="table3" cellpadding="0" style="border-collapse: collapse">
		<tr>
			<td width="210" valign="top">
<table border="0" width="210" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td valign="top">
		<table border="0" width="210" id="table1" cellpadding="0" style="border-collapse: collapse">
			<tr>
				<td height="19" width="12">
				<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
				<td height="19" background="tema/images/blok_2.gif" width="960">
				<p align="center">
				<font class="baslik">
				Menü</font></td>
				<td height="19" width="10">
				<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
			</tr>
			<tr>
				<td colspan="3" class="blok_bg">
<div align="center">
	<table border="0" width="90%" id="table1" cellpadding="0" style="border-collapse: collapse">
<%
set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from kategori"
rs.open SQL,data,1,3
if not rs.eof then
%>
		<tr>
			<td width="20"><img src="tema/images/mini-category.gif"></td>
			<td><font class="blok"><b>Kategoriler</b></font></td>
		</tr>
<%
Do While Not rs.EOF

set blog = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as blog_say from blog where kat_id= "&rs("id")&""
blog.open SQL,data,1,3
%>
		<tr>
			<td width="20"></td>
			<td><img src="tema/images/kategori.gif"> <a href="kategori.asp?id=<%=rs("id")%>" title="<%=rs("aciklama")%>"><%=rs("ad")%></a> <font class="blok">(<%=blog("blog_say")%>)</font></td>
		</tr>
<%
blog.close
set blog = Nothing

rs.MoveNext
Loop
end if
rs.Close
Set rs = Nothing
%>
<%
set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from linkler"
rs.open SQL,data,1,3
if not rs.eof then
%>
		<tr>
			<td width="20"><img src="tema/images/linkler.gif"></td>
			<td><font class="blok"><b>Bağlantılar</b></font></td>
		</tr>
<% Do While not rs.eof %>
		<tr>
			<td width="20"></td>
			<td><img src="tema/images/link.gif"> <a href="<%=rs("link")%>" target="_blank"><%=rs("isim")%></a></td>
		</tr>
<%
rs.MoveNext
Loop
End if
rs.Close
Set rs = Nothing
%>
	</table>
</div>
      			</td>
			</tr>
			<tr>
				<td colspan="3" height="14">
				<img border="0" src="tema/images/blok_4.gif" width="210" height="14"></td>
			</tr>
		</table>
		</td>
	</tr>
<%
set ab = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from ankets order by id desc"
ab.open SQL,data,1,3
if not ab.eof then
%>
	<tr>
		<td height="12"></td>
	</tr>
	<table border="0" width="210" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td valign="top">
		<table border="0" width="210" id="table1" cellpadding="0" style="border-collapse: collapse">
			<tr>
				<td height="19" width="12">
				<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
				<td height="19" background="tema/images/blok_2.gif" width="960">
				<p align="center">
				<font class="baslik">Anket</font></td>
				<td height="19" width="10">
				<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
			</tr>
			<tr>
				<td colspan="3" class="blok_bg">
<div align="center">
	<table border="0" width="90%" id="table1" cellpadding="0" style="border-collapse: collapse">
<form action="islem.asp?islem=anket&id=<%=ab("id")%>&s=<%=session.sessionID%>" method="post">
		<tr>
			<td width="100%" colspan="2" align="center">
			<b><font class="blok"><%=ab("soru")%></font></b></td>
		</tr>
<!--Anket Şıkları-->
<%
        Set tAnc = Server.CreateObject("ADODB.RecordSet") 
        tAnc.Open "anket Where a_id Like '"&ab("id")&"'",data,1,3
        Vote = 0
        Do While Not tAnc.EOF
        Vote = Vote + tAnc("deger")
        tAnc.MoveNext
        Loop
        tAnc.Close
        Set tAnc = NoThing

        Set tAnc = Server.CreateObject("ADODB.RecordSet")
        tAnc.Open "anket Where a_id Like '"&ab("id")&"'",data,1,3
Do While Not tAnc.EOF

                strOy = tAnc("deger")
                If strOy = "0" Then
                tOy = "0"
                Else
                tOy = (strOy /Vote) * 100
                End If
%>
		<tr>
			<td width="10%"><input type="radio" name="cevap" value="<%=tAnc("id")%>"></td>
			<td width="90%"><font class="blok"><%=tAnc("cevap")%> (%<%=Left(tOy,4)%>)</font></td>
		</tr>
		<tr>
			<td width="10%"></td>
			<td width="90%">
<div style="width: 100px; height: 5px; border: 1px solid #000000">
<img src="tema/images/vote.gif" height="5" width="<%=Int(tOy)%>"></div>
			</td>
		</tr>
<% tAnc.MoveNext
                Loop 
                Set tOy = data.Execute("anket Where a_id Like '"&ab("id")&"'")
                Do While Not tOy.EOF
                AraToplam = tOy("deger")
                OySayisi = OySayisi + AraToplam
                tOy.MoveNext
                Loop
                tOy.Close : Set tOy = Nothing	%>

		<tr>
			<td width="100%" colspan="2">
			<p align="center"><font class="blok">Toplam Oy: <%=OySayisi%></font></p>
			</td>
		</tr>
		<tr>
			<td width="100%" colspan="2">
			<p align="center"><input type="submit" value="Oy Ver" class="dugme">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="center"><a href="anketler.asp">Tüm Anketler</a></td>
		</tr>
</form>
<!--/Anket Şıkları-->
	</table>
</div>
      			</td>
			</tr>
			<tr>
				<td colspan="3" height="14">
				<img border="0" src="tema/images/blok_4.gif" width="210" height="14"></td>
			</tr>
		</table>
		</td>
	</tr>
<%
end if
ab.Close
Set ab = Nothing
%>
	<tr>
		<td height="12"></td>
	</tr>
	<table border="0" width="210" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td valign="top">
		<table border="0" width="210" id="table1" cellpadding="0" style="border-collapse: collapse">
			<tr>
				<td height="19" width="12">
				<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
				<td height="19" background="tema/images/blok_2.gif" width="960">
				<p align="center">
				<font class="baslik">
				Takvim</font></td>
				<td height="19" width="10">
				<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
			</tr>
			<tr>
				<td colspan="3" class="blok_bg">
	<%
	Function temizle(trh)
	if trh<10 then
		trh=Right(trh,1)
	end if
	temizle=trh
	End Function
	
	d_month=Request.QueryString("month")
	d_year=Request.QueryString("year")
	if d_year="" or d_year<1900 or d_year>2100 then
	d_year=Year(date)
	end if
	if d_month="" then d_month=month(date)
	if d_year="" then d_year=year(date)
	if d_month>12 then
	d_month=1
	d_year=d_year+1
	end if
	if d_month<1 then
	d_month=12
	d_year=d_year-1
	end if
	d_day=28
	while isdate(d_day&","&d_month&","&d_year)=true
	d_day=d_day+1
	wend
	d_day=d_day-1
	%>
	<table border="0" cellspacing="0" cellpadding="0" width="90%" align="center" style="border-collapse: collapse">
	<tr>
		<td><a href="index.asp?month=<%=d_month-1%>&year=<%=d_year%>">«</a></td>
		<td align=center colspan=5><font class="blok"><b><%=Monthname(d_month)%> - <%=d_year%></b></font></td>
		<td>
		<p align="right"><a href="index.asp?month=<%=d_month+1%>&year=<%=d_year%>">»</a></td>
	</tr>
	<tr>
	<td class="blok" align="center"><b>PT</b></font></td>
	<td class="blok" align="center"><b>SL</b></font></td>
	<td class="blok" align="center"><b>ÇŞ</b></font></td>
	<td class="blok" align="center"><b>PŞ</b></font></td>
	<td class="blok" align="center"><b>CM</b></font></td>
	<td class="blok" align="center"><b>CT</b></font></td>
	<td class="blok" align="center"><b>PZ</b></font></td>
	</tr>
	<tr>
	<%
	sayac=weekday("01,"&d_month&","&d_year)
	if sayac=1 then 
		sayac=6
	else
		sayac=sayac-2
	end if	
	
	for w=1 to sayac
	Response.Write("<td> </td>")
	next
	
	for w=1 to d_day
	if w = Day(Now()) then
	a = w
	else
	a = w
	end if
	if sayac mod 7=0 then Response.Write("</tr> <tr>")
	itarih=Cint(d_month)&"/"&a&"/"&Cint(d_year)
	
	set rs = Server.CreateObject("ADODB.RecordSet")
	SQL = "select tarih,konu from blog where tarih like '%"&itarih&"%'"
	rs.open SQL,data,1,3
	if rs.eof then
	%>
	<td align="center"<% IF a = Day(Now()) and Cint(d_month) = month(date) and Cint(d_year) = year(date) Then%> class="takvimbugun"<% End IF %>><font class="takvimpasif"><%=a%></font></td>
	<%Else%>
	<td align="center"<% IF a = Day(Now()) and Cint(d_month) = month(date) and Cint(d_year) = year(date) Then%> class="takvimbugun"<% else %> style="background-color:#FFFFEC;"<%end if%>><font class="blok"><a href="takvim.asp?gun=<%=a%>&ay=<%=d_month%>&yil=<%=d_year%>" title="<%=rs("konu")%>"><b><%=a%></b></a></font></td>
	<%End if : rs.close : set rs=nothing%>
	<%
	sayac=sayac+1
	next
	%>
	</tr>
	</table>
	<div align="center">
<table border="0" id="table1" cellpadding="0" style="border-collapse: collapse">
		<form method="GET" action="index.asp">
	<tr>
		<td>
			<select name="month" class="alan" size="1">
			<option value="1">Ocak</option>
			<option value="2">Şubat</option>
			<option value="3">Mart</option>
			<option value="4">Nisan</option>
			<option value="5">Mayıs</option>
			<option value="6">Haziran</option>
			<option value="7">Temmuz</option>
			<option value="8">Ağustos</option>
			<option value="9">Eylül</option>
			<option value="10">Ekim</option>
			<option value="11">Kasım</option>
			<option value="12">Aralık</option>
			</select>
		</td>
		<td>
		<input type="submit" value="Git" align="center" class="dugme" onClick="this.form.submit();this.disabled=true; return true;"></td>
	</tr>
		</form>
</table>
	</div>
	</td>
      			</td>
			</tr>
			<tr>
				<td colspan="3" height="14">
				<img border="0" src="tema/images/blok_4.gif" width="210" height="14"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td height="12"></td>
	</tr>
	<tr>
		<td valign="top">
		<table border="0" width="210" id="table1" cellpadding="0" style="border-collapse: collapse">
			<tr>
				<td height="19" width="12">
				<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
				<td height="19" background="tema/images/blok_2.gif" width="960">
				<p align="center">
				<font class="baslik">
				İstatistikler</font></td>
				<td height="19" width="10">
				<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
			</tr>
			<tr>
				<td colspan="3" class="blok_bg">
				<%
Set rs=Server.CreateObject("Adodb.Recordset")
sql = "Select * from sayac " 
rs.open sql, data, 1, 3
%>

<%
if Request.Cookies("saydir")("hit") = "kapat" then
else
rs("sayac")=rs("sayac") + 1
rs.update
   Response.Cookies("saydir")("hit") = "kapat"
   Response.Cookies("saydir").Expires = Now() + 1
end if
%>
				<div align="center">
<table border="0" width="90%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td height="14">
		<font class="blok">
		&nbsp;Toplam Hit: <%=rs("sayac")%></font></td>
	</tr>
	<tr>
		<td height="14">
		<font class="blok">
		&nbsp;Sitede Aktif: <!-- #include file="../aktif.asp" --></font></td>
	</tr>
	<tr>
		<td height="14">
		<font class="blok">
		&nbsp;Ip: <% Response.write ""&Request.ServerVariables("REMOTE_ADDR")&"" %></font></td>
	</tr>
	<tr>
		<td height="14">
		<font class="blok">
		&nbsp;Browser: <%=Server.CreateObject("MSWC.BrowserType" ).Browser %> - <%=Server.CreateObject("MSWC.BrowserType" ).Version %></font></td>
	</tr>
<%
set kat = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as kat_say from kategori"
kat.open SQL,data,1,3
%>
	<tr>
		<td height="14">
				<font class="blok">
&nbsp;Toplam Kategori: <%=kat("kat_say")%></font></td>
	</tr>
<%
kat.close
set kat = Nothing
%>
<%
set blg = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as blg_say from blog"
blg.open SQL,data,1,3
%>
	<tr>
		<td height="14">
				<font class="blok">
&nbsp;Toplam Blog: <%=blg("blg_say")%></font></td>
	</tr>
<%
blg.close
set blg = Nothing
%>
<%
set yrm = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as yrm_say from yorum where onay=0"
yrm.open SQL,data,1,3
%>
	<tr>
		<td height="14">
				<font class="blok">
&nbsp;Toplam Yorum: <%=yrm("yrm_say")%></font></td>
	</tr>
<%
yrm.close
set yrm = Nothing
%>
<%
set rsm = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as rsm_say from galeri"
rsm.open SQL,data,1,3
%>
	<tr>
		<td height="14">
				<font class="blok">
&nbsp;Toplam Resim: <%=rsm("rsm_say")%></font></td>
	</tr>
<%
rsm.close
set rsm = Nothing
%>
<%
set zd = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as zd_say from zd where onay=0"
zd.open SQL,data,1,3
%>
	<tr>
		<td height="14">
				<font class="blok">
&nbsp;Toplam Mesaj: <%=zd("zd_say")%></font></td>
	</tr>
<%
zd.close
set zd = Nothing
%>
</table></div>
				</td>
			</tr>
			<tr>
				<td colspan="3" height="14">
				<img border="0" src="tema/images/blok_4.gif" width="210" height="14"></td>
			</tr>
		</table>
      </td>
	</tr>
	<tr>
		<td height="12"></td>
	</tr>
	<tr>
		<td valign="top">
		<table border="0" width="210" id="table1" cellpadding="0" style="border-collapse: collapse">
			<tr>
				<td height="19" width="12">
				<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
				<td height="19" background="tema/images/blok_2.gif" width="960">
				<p align="center">
				<font class="baslik">Etiket Bulutu</font></td>
				<td height="19" width="10">
				<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
			</tr>
			<tr>
				<td colspan="3" class="blok_bg">
<table border="0" width="90%" id="table1" cellpadding="0" style="border-collapse: collapse" align="center">
	<tr>
		<td align="center">
<%
set objetiket=data.execute("SELECT DISTINCT etiket FROM etiket")
if not Objetiket.eof then
renk="0"
Do While Not Objetiket.eof
renk=renk+1

if renk=1 then color="blue"
if renk=2 then color="red"
if renk=3 then color="green"
if renk=4 then color="gray"
if renk=5 then color="black"

	Set ObjBulut = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT Count(Etiket) as TagHit FROM etiket WHERE Etiket='" & Objetiket("Etiket") & "'"
	ObjBulut.Open SQL, data, 3, 3
	hit=ObjBulut("TagHit")
	boyut=Int(10 + hit)
	%><a href="etiket.asp?etiket=<%=Objetiket("Etiket")%>"><span title="<%=hit%> Kayıt" style="font-size:<%=boyut%>px; color:<%=color%>;"><%=Objetiket("Etiket")%></span></a> <%
Objetiket.MoveNext
if renk>4 then renk=0
Loop
end if
Objetiket.close : set objetiket=nothing
%>
		</td>
	</tr>
</table>
				</td>
			</tr>
			<tr>
				<td colspan="3" height="14">
				<img border="0" src="tema/images/blok_4.gif" width="210" height="14"></td>
			</tr>
		</table>
      </td>
	</tr>
</table>
</td>

			<td width="26" rowspan="3">
			<img border="0" src="tema/images/spacer.gif" width="28" height="1"></td>
			
			<td width="550" rowspan="3" valign="top">
<% Call Govde %>
			</td>
		</tr>
	</table>
</div><br>
			</td>
		</tr>
		<tr>
			<td colspan="3" align="center" bgcolor="#ffffff" background="tema/v3_img/govde_orta.gif">
<script type="text/javascript"><!--
google_ad_client = "pub-2790849200120182";
/* 728x90, oluşturulma 12.06.2008 */
google_ad_slot = "5251880239";
google_ad_width = 728;
google_ad_height = 90;
//-->
</script>
<script type="text/javascript"
src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
</script>
			</td>
		</tr>
		<tr>
			<td width="19" background="tema/v3_img/v3_22.gif"></td>
			<td width="784" bgcolor="#BDE4EA" valign="top">
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="33%"><font class="b_baslik">En Çok Okunanlar</font></td>
		<td width="33%"><font class="b_baslik">Son Yorumlananlar</font></td>
		<td width="33%"><font class="b_baslik">Hakkımda</font></td>
	</tr>
	<tr>
		<td width="33%" valign="top"><%
set blog = Server.CreateObject("ADODB.RecordSet")
SQL = "select id,konu,hit from blog order by hit DESC"
blog.open SQL,data,1,3

For p = 1 To 10
if blog.eof Then exit For
%><a href="<%=SEOLink(blog("id"))%>"><%=blog("konu")%></a> (<font class="orta"><%=blog("hit")%></font>)<br><%
blog.movenext
Next
blog.Close
Set blog = Nothing
%>
		</td>
		<td width="33%" valign="top"><%
set yorum = Server.CreateObject("ADODB.RecordSet")
SQL = "SELECT blog_id FROM yorum where onay=0 and blog_id<>0 GROUP BY blog_id order by Max(tarih) desc"
yorum.open SQL,data,1,3
if not yorum.eof then

For p = 1 To 10
if yorum.eof Then exit For

set blog = Server.CreateObject("ADODB.RecordSet")
SQL = "select konu,id from blog where id="&yorum("blog_id")&""
blog.open SQL,data,1,3

set syorum = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as yorum_say from yorum where blog_id= "&blog("id")&" and onay=0"
syorum.open SQL,data,1,3
%><a href="<%=SEOLink(blog("id"))%>#yorumlar"><%=blog("konu")%></a> (<font class="orta"><%=syorum("yorum_say")%></font>)<br><%
syorum.close
set syorum = Nothing

blog.close : set blog = nothing
yorum.Movenext 
Next

Else
response.write "<center><b><font class=""orta"">Yorumlanan Konu Yok</font></b></center>"
end if
yorum.close : set yorum=nothing

if not hakkimda="" then
	hakkimdayazi=Left(Cevir(hakkimda),400)
else
	hakkimdayazi=""
end if
%>		</td>
		<td width="33%" valign="top"><a href="hakkimda.asp"><%=hakkimdayazi%></a></td>
		</tr>
</table>
			</td>
			<td width="27" style="background-image: url('tema/v3_img/v3_23.gif'); background-repeat: repeat-y; background-position-x: right"></td>
		</tr>
	</table>
</div>
<div align="center">
	<table border="0" width="830" id="table4" cellpadding="0" style="border-collapse: collapse">
		<tr>
			<td width="35" rowspan="2">
			<img border="0" src="tema/v3_img/v3_24.gif" width="145" height="57"></td>
			<td width="686" height="31" bgcolor="#BDE4EA" background="tema/v3_img/v3_25.gif"></td>
			<td width="27" height="31">
			<a href="javascript:history.back()"><img border="0" src="tema/v3_img/v3_26.gif" width="33" height="31" border="0"></a></td>
			<td width="25" height="31">
			<a href="javascript:window.location.reload()"><img border="0" src="tema/v3_img/v3_27.gif" width="33" height="31" border="0"></a></td>
			<td width="30" height="31">
			<a href="javascript:history.forward(1)"><img border="0" src="tema/v3_img/v3_28.gif" width="31" height="31" border="0"></a></td>
			<td width="27" height="31">
			<img border="0" src="tema/v3_img/v3_29.gif" width="27" height="31"></td>
		</tr>
		<tr>
			<td width="658" colspan="4" background="tema/v3_img/v3_30.gif">
			</td>
			<td width="27">
			<img border="0" src="tema/v3_img/v3_31.gif" width="27" height="26"></td>
		</tr>
		<tr>
			<td width="830" colspan="6">
			<img border="0" src="tema/v3_img/v3_32.gif" width="830" height="18"></td>
		</tr>
	</table>
</div>

</body>
<%
data.close : set data=nothing
%>
</html>