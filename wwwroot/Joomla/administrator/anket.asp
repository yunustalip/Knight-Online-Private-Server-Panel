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
%><head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" />
<meta name="keywords" content="JoomlASP, Joomla, MySQL, ASP, Active Server Page, ASP Portal, JoomlASP temalarý, JoomlASP modülleri, JoomlASP bileþenleri, Site içerik yönetimi, JoomlASP Portalý">
<meta name="description" content="JoomlASP - Geliþime Açýk Site Ýçerik Yönetimi">
<meta name="author" content="JoomlASP | Hasan Emre Asker">
<title>JoomlASP Site Yönetici Paneli v1.2</title>
<link href="favicon.ico" rel="JoomlASP" />
<link href="admin.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.style6 {font-size: 24px}
-->
</style>
</head>
<!--#include file="kontrol.asp"-->
<%
function mysqltarihsaathim(varDate)
if day(varDate) < 10 then
dd = "0" & day(varDate)
else
dd = day(varDate)
end if

if month(varDate) < 10 then
mm = "0" & month(varDate)
else
mm = month(varDate)
end if

if hour(varDate) < 10 then
hh = "0" & hour(varDate)
else
hh = hour(varDate)
end if

if minute(varDate) < 10 then
mi = "0" & minute(varDate)
else
mi = minute(varDate)
end if

if second(varDate) < 10 then
se = "0" & second(varDate)
else
se = second(varDate)
end if

mysqltarihsaathim = Year(varDate)&"-"&mm&"-"&dd
end function
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr>
    <td colspan="2" valign="top" background="../images/admin_top.png"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../images/admin_banner.png" width="307" height="36" /></td>
          <td width="58"><img src="../images/admin_banner_son.png" width="58" height="36" /></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="150" valign="top"><!--#include file="menu.asp"--></td>
    <td valign="top"><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
      <tr>
        <!--#include file="haberler.asp"-->
      </tr>
      <tr>
        <td colspan="3"><table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/anket.png" width="128" height="128" align="middle" /><span class="style6"> Anket Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr bgcolor="e58e4d">
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td bgcolor="#333333" class="style4">
<%
dim SQL, rs, baglanti, id

id = Request.QueryString("id")


if Request.QueryString("sub") = "addnew" then addnew()
if Request.QueryString("sub") = "edit" then edit()
if Request.QueryString("sub") = "edit_add" then edit_add()

SQL = "SELECT * FROM gop_anketsoru ORDER BY id ASC"
set rs = baglanti.Execute(SQL)
%>
<%if 1 = 1 then%>

<div align="center">
<table border="1" class="nortxtv8" cellpadding="3" cellspacing="0" width="600" style="border-collapse: collapse" bordercolor="#000000">
	<tr>
		<td colspan="2" bgcolor="#ADD8E6">
		<a href="anket.asp?sub=addnew"><img src="adm_img/new.gif" name="new" border="0" alt="Add new poll" WIDTH="61" HEIGHT="16" align="absmiddle"></a></td>
	</tr>
	<tr>
		<%
		if not rs.eof then
		do
		%>
		<tr>
		<td width="425"><%="Poll id: <font color=""#FF0000""><b>" & rs("id") & "</b></font> - " & rs("title")%></td>
		<td width="157" align="center">
		<%if rs("active") then %>
		<a href="anket_guncelle.asp?sub=inact&id=<%=rs("id")%>"><img src="adm_img/active.gif" align="absmiddle" name="active" border="0" alt="The poll is active - set to inactive" WIDTH="36" HEIGHT="16"></a>
		<%else%>
		<a href="anket_guncelle.asp?sub=act&id=<%=rs("id")%>"><img src="adm_img/inact.gif" align="absmiddle" name="inact" border="0" alt="The poll is inactive - set to active" WIDTH="36" HEIGHT="16"></a>
		<%end if%>
		<a href="anket.asp?sub=edit&id=<%=rs("id")%>"><img src="adm_img/edit.gif" align="absmiddle" name="edit" border="0" alt="Edit poll" WIDTH="50" HEIGHT="16"></a>
		<a href="anket_guncelle.asp?sub=del&id=<%=rs("id")%>"><img src="adm_img/delete.gif" align="absmiddle" name="delete" border="0" alt="Delete poll" WIDTH="36" HEIGHT="16"></a></td>
		</tr>
		<%
		rs.movenext
		loop until rs.eof
		else
		%>
			  <td>Anket Yok </a>
		          <%end if%>
	</tr>
</table>
</div>
	
<%
elseif 2 = 2 then

	addnew()
	
end if
%>

<%sub addnew()%>

<div align="center">
<table border="1" class="nortxtv8" cellpadding="3" cellspacing="0" width="600" style="border-collapse: collapse" bordercolor="#000000">
	<tr>
		<td colspan="2" bgcolor="#ADD8E6"><b>Add new poll</b> </td>
	</tr>
	<form name="formPoll" method="post" action="anket_guncelle.asp?sub=new">
	<tr>
		<td>
			<table width="100%" border="0" cellspacing="0" cellpadding="1" class="nortxtv8">
			  <tr>
			    <td width="100">Anket Sorusu </td>
			    <td><input type="text" name="title" size="40" class="nortxtv8"></td>
			  </tr>
			  <tr>
			    <td width="100">Cevap 1</td>
			    <td><input type="text" name="a1" size="40" class="nortxtv8"></td>
			  </tr>
			  <tr>
			    <td width="100">Cevap 2</td>
			    <td><input type="text" name="a2" size="40" class="nortxtv8"></td>
			  </tr>
			  <tr>
			    <td width="100">Cevap 3</td>
			    <td><input type="text" name="a3" size="40" class="nortxtv8"></td>
			  </tr>
			  <tr>
			    <td width="100">Cevap 4</td>
			    <td><input type="text" name="a4" size="40" class="nortxtv8"></td>
			  </tr>
			  <tr>
			    <td width="100">Cevap 5</td>
			    <td><input type="text" name="a5" size="40" class="nortxtv8"></td>
			  </tr>
			  <tr>
			    <td width="100">Cevap 6</td>
			    <td><input type="text" name="a6" size="40" class="nortxtv8"></td>
			  </tr>
			  <tr>
			    <td width="100">Cevap 7</td>
			    <td><input type="text" name="a7" size="40" class="nortxtv8"></td>
			  </tr>
			  <tr>
			    <td width="100">Cevap 8</td>
			    <td><input type="text" name="a8" size="40" class="nortxtv8"></td>
			  </tr>
			  <tr>
			    <td width="100">Cevap 9</td>
			    <td><input type="text" name="a9" size="40" class="nortxtv8"></td>
			  </tr>
			  <tr>
			    <td width="100">Cevap 10</td>
			    <td><input type="text" name="a10" size="40" class="nortxtv8"></td>
			  </tr>
			  <tr>
			    <td width="100">Baþlama Tarihi </td>
			    <td><input type="text" name="d_s" size="20" class="nortxtv8"></td>
			  </tr>
			  <tr>
			    <td width="100">Bitiþ Tarihi </td>
			    <td><input type="text" name="d_e" size="20" class="nortxtv8"></td>
			  </tr>
			  <tr>
			    <td width="100">&nbsp;</td>
			    <td><input type="image" border="0" name="Submit" src="adm_img/submit.gif" width="40" height="16"></td>
			  </tr>
			</table>
		</td>
	</tr>
	</form>
</table>
</div>

<%end sub%>

<%
sub edit()
dim an_no, show
dim e_start, e_end
show = 1

an_no = 1

SQL = "SELECT * FROM gop_anketsoru, gop_anketcevap WHERE id=" & id & " AND id=poll_id"
set rs = baglanti.Execute(SQL)

'if there is no answers in the database
if rs.eof then
	SQL = "SELECT * FROM gop_anketsoru WHERE id=" & id
	set rs = baglanti.Execute(SQL)
	show = 0
end if

e_start = mysqltarihsaathim(rs("expiration_start"))
e_end = mysqltarihsaathim(rs("expiration_end"))

%>

<div align="center">
<table border="1" class="nortxtv8" cellpadding="3" cellspacing="0" width="600" style="border-collapse: collapse" bordercolor="#000000">
	<tr>
		<td colspan="2" bgcolor="#ADD8E6"><b>Anket Düzenleme </b> </td>
	</tr>
	<form name="formPoll" method="post" action="anket_guncelle.asp?sub=edit&id=<%=id%>">
	<tr>
		<td>
			<table width="100%" border="0" cellspacing="0" cellpadding="1" class="nortxtv8">
			  <tr>
			    <td width="100">Anket Sorusu</td>
			    <td><input type="text" name="title" size="40" class="nortxtv8" value="<%=rs("title")%>">&nbsp;Anket id: <font color="#FF0000"><b><%=rs("id")%></b></font>
			    &nbsp;&nbsp;Toplam Oy: <b><%=rs("votes")%></b></td>
			  </tr>
			  <%
			  if not show = 0 then
			  do
			  %>			  
			  <tr>
			    <td width="100">Cevap <%=an_no%></td>
			    <td>
			    <input type="hidden" name="h<%=an_no%>" size="40" class="nortxtv8" value="<%=rs("answer_id")%>">
			    <input type="text" name="a<%=an_no%>" size="40" class="nortxtv8" value="<%=rs("answer")%>">&nbsp;Cevap id: <font color="#FF0000"><b><%=rs("answer_id")%></b></font>
			    &nbsp;&nbsp;Oy: <b><%=rs("no_votes")%></b>&nbsp;
			    <a href="anket_guncelle.asp?sub=del_answ&answ_id=<%=rs("answer_id")%>&id=<%=rs("id")%>">
			    <img src="adm_img/delete.gif" align="absmiddle" name="delete" border="0" alt="Delete poll" WIDTH="36" HEIGHT="16">
			    </a></td>
			    </td>
			  </tr>
			  <%
			  an_no = an_no + 1
			  rs.movenext
			  loop until rs.eof
			  end if
			  %>
			  <tr>
			    <td width="100">Baþlama Tarihi</td>
			    <td><input type="text" name="d_s" size="20" class="nortxtv8" value="<%=e_start%>"></td>
			  </tr>
			  <tr>
			    <td width="100">Bitiþ Tarihi</td>
			    <td><input type="text" name="d_e" size="20" class="nortxtv8" value="<%=e_end%>"></td>
			  </tr>
			  <tr>
			    <td width="100">&nbsp;</td>
			    <td>
			    <input type="hidden" name="no_answers" size="40" class="nortxtv8" value="<%=an_no - 1%>">
			    <input type="image" border="0" name="Submit" src="adm_img/submit.gif" width="40" height="16" alt="Update database">
			    <a href="anket.asp?sub=edit_add&id=<%=id%>"><img border="0" name="add" src="adm_img/add_answ.gif" width="63" height="16" alt="Add one answer"></a>
			    </td>
			  </tr>
			</table>
		</td>
	</tr>
	</form>
</table>
</div>

<%
end sub
%>

<%
sub edit_add()
dim an_no, show
show = 1

an_no = 1

SQL = "SELECT * FROM gop_anketsoru, gop_anketcevap WHERE id=" & id & " AND id=poll_id"
set rs = baglanti.Execute(SQL)

'if there is no answers in the database
if rs.eof then
	SQL = "SELECT * FROM gop_anketsoru WHERE id=" & id
	set rs = baglanti.Execute(SQL)
	show = 0
end if



e_start = rs("expiration_start")
e_end = rs("expiration_end")
%>

<div align="center">
<table border="1" class="nortxtv8" cellpadding="3" cellspacing="0" width="600" style="border-collapse: collapse" bordercolor="#000000">
	<tr>
		<td colspan="2" bgcolor="#ADD8E6"><b>Anket Düzenleme </b> </td>
	</tr>
	<form name="formPoll" method="post" action="anket_guncelle.asp?sub=edit_add&id=<%=id%>">
	<tr>
		<td>
			<table width="100%" border="0" cellspacing="0" cellpadding="1" class="nortxtv8">
			  <tr>
			    <td width="100">Anket Sorusu </td>
			    <td><b><%=rs("title")%></b>&nbsp;&nbsp;&nbsp;Anket id: <font color="#FF0000"><b><%=rs("id")%></b></font>
			    &nbsp;&nbsp;Toplam Oy: <b><%=rs("votes")%></b></td>
			  </tr>
			  <%
			  if not show = 0 then
			  do
			  %>			  
			  <tr>
			    <td width="100">Cevap <%=an_no%></td>
			    <td><b><%=rs("answer")%></b>&nbsp;&nbsp;&nbsp;Cevap id: <font color="#FF0000"><b><%=rs("answer_id")%></b></font>
			    &nbsp;&nbsp;Oy: <b><%=rs("no_votes")%></b>
			    </td>
			  </tr>
			  <%
			  an_no = an_no + 1
			  rs.movenext
			  loop until rs.eof
			  end if
			  %>
			  <tr>
				<td width="100">Yeni Cevap </td>
				<td>
					<input type="text" name="add_one" size="40" class="nortxtv8">
				</td>
			  </tr>
			  <tr>
			    <td width="100">&nbsp;</td>
			    <td>
			    <input type="hidden" name="no_answers" size="40" class="nortxtv8" value="<%=an_no - 1%>">
			    <input type="image" border="0" name="Submit" src="adm_img/submit.gif" width="40" height="16" alt="Update database">
			    </td>
			  </tr>
			</table>
		</td>
	</tr>
	</form>
</table>
</div>

<%
end sub
%>
</td>
                  </tr>

              </table>
</td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="20" colspan="3" bgcolor="#CC0000"><strong><span class="style4">&nbsp;JoomlASP Admin Panel Bilgisi </span></strong></td>
      </tr>
      <tr>
        <td colspan="3" bgcolor="#CCCCCC"><p>Web sitenizi en iyi hale getirebilmek için sol bölümde bulunan linkleri kullanarak sitenize Yeni Linkler, Kategoriler, Alt Kategoriler ve Veriler girebilir, Üyelerinizi düzenleyip bilgi deðiþikliði iþlemlerini gerçekleþtirebilirsiniz.</p>
            <p>Sisteminizi daha iyi ve kararlý hale getirmek için lütfen <strong>JoomlASP Resmi Destek Sitesi</strong> (RDS) olan www.joomlasp.com adresindeki <strong>Forum</strong> bölümünü kullanýnýz. Baþka sitelerden alýnan bileþen ve modüllerden doðacak sorunlardan JoomlASP sorumlu deðildir.</p></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="25" colspan="2" background="../images/admin_top2.png"><!--#include file="alt.asp"--></td>
  </tr>
</table>
