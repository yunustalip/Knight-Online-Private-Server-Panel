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
'		Tel : 0544 275 9804
'		Mail: emre06@hotmail.com.tr / info@joomlasp.com/.net/.org
'
'
'		Lisans Anlaþmasý Gereði Lütfen Google Reklam Bölümünü Sitenizden kaldýrmayýnýz. Bu sizin GOOGLE reklamlarýný yapmanýza
'		kesinlikle bir engel deðildir. reklam.asp içeriðinin yada yayýnladýðý verinin deðiþmesi lisans politikasýnýn dýþýna çýkýlmasýna
'		ve JoomlASP CMS sistemini ücretsiz yayýnlamak yerine ücretlie hale getirmeye bizi teþfik etmektedir. Bu Sistem için verilen emeðe
'		saygý ve bir çeþit ödeme seçeneði olarak GOOGLE reklamýmýzýn deðiþtirmemesi yada silinmemesi gerekmektedir.
%>
<!--#include file="kontrol.asp"-->
<%
dim SQL, baglanti, rs

if Request.QueryString("sub") = "new" then add_new()
if Request.QueryString("sub") = "act" then act()
if Request.QueryString("sub") = "inact" then inact()
if Request.QueryString("sub") = "edit" then edit()
if Request.QueryString("sub") = "edit_add" then edit_add()
if Request.QueryString("sub") = "del_answ" then del_answ()
if Request.QueryString("sub") = "del" then del()
%>


<%
'add new poll
sub add_new()
dim a, b, poll_id, answ_id
dim email, str

'get last poll id
SQL = "SELECT * FROM gop_anketsoru ORDER BY id DESC"
set rs = baglanti.Execute(SQL)

if not rs.eof then
	poll_id = rs("id") + 1
else
	poll_id = 1
end if

'add poll title
SQL = "INSERT INTO gop_anketsoru (id,title,expiration_start,expiration_end) VALUES " & _
	  "(" & poll_id & ",'" & Request.Form("title") & "','" & Request.Form("d_s") & "','" & Request.Form("d_e") & "')"
set rs = baglanti.Execute(SQL)

'get last answer id
SQL = "SELECT * FROM gop_anketcevap ORDER BY answer_id DESC"
set rs = baglanti.Execute(SQL)

if not rs.eof then
	answ_id = rs("answer_id")
else
	answ_id = 1
end if

'add answers
a = 1 'for looping through textboxes
b = 1 'for adding answer_id
do

if not Request.Form("a" & a) = "" then
	SQL = "INSERT INTO gop_anketcevap (poll_id,answer_id,answer) VALUES " & _
		  "(" & poll_id & "," & answ_id + b & ",'" & Request.Form("a" & a) & "')"
	set rs = baglanti.Execute(SQL)
	b = b + 1
end if

a = a + 1
loop until a = 11

'code for sending email to admin users (only admin can set poll to active)
'######### UNCOMENT FOR SENDING EMAIL TO ALL ADMIN USERS WHEN NORMAL USER ADDS NEW POLL #########
'if session("adminRights") = 2 then
	
	'SQL = "SELECT * FROM users WHERE status = 1 ORDER BY usr_id ASC"
	'set rs = baglanti.Execute(SQL)
	
	'do
	
	'if not email = "" then 
	'	str = "; " & email 
	'end if

	'email = rs("email") & str 

	'rs.movenext
	'loop until rs.eof

	'Set EMail = CreateObject("CDONTS.NewMail") 
	'With EMail
	' .From = "someone@domain.com"
	' .bcc = email 
	' .Subject = "New poll in database"
	' .Body = "Poll title: " & Request.Form("title")
	' .Send 
	'End with 
	'Set EMail = Nothing 

'end if

Response.Redirect("anket.asp")

end sub
%>

<%
'set poll to active
sub act()
dim poll_id

poll_id = Request.QueryString("id")

'set all polls to inactive
SQL = "UPDATE gop_anketsoru SET active=False"
set rs = baglanti.execute(SQL)

'set selected poll to active
SQL = "UPDATE gop_anketsoru SET active=True WHERE id=" & poll_id
set rs = baglanti.execute(SQL)

Response.Redirect("anket.asp")

end sub
%>

<%
'set poll to inactive
sub inact()
dim poll_id

poll_id = Request.QueryString("id")

'set selected poll to inactive
SQL = "UPDATE gop_anketsoru SET active=False WHERE id=" & poll_id
set rs = baglanti.execute(SQL)

Response.Redirect("anket.asp")

end sub
%>

<%
'save edited poll
sub edit()
dim no, i, poll_id

i = 1

'get number of textboxes
no = Request.Form ("no_answers")
'get poll id
poll_id = Request.QueryString ("id")

'update poll title
SQL = "UPDATE gop_anketsoru SET title='" & Request.Form ("title") & "', expiration_start='" & Request.Form("d_s") & _
	  "', expiration_end='" & Request.Form("d_e") & "' WHERE id=" & poll_id
set rs = baglanti.execute(SQL)

'get all textboxes and update database with edited answers
do
SQL = "UPDATE gop_anketcevap SET answer='" & Request.Form ("a" & i) & "' WHERE answer_id=" & Request.Form ("h" & i)
set rs = baglanti.execute(SQL)
i = i + 1
loop until i = no + 1

Response.Redirect("anket.asp?sub=edit&id=" & poll_id)

end sub
%>

<%
'add one answer
sub edit_add()
dim poll_id, answ_id

poll_id = Request.QueryString("id")

'if textbox is empty ...
if not Request.Form("add_one") = "" then

'get last answer id
SQL = "SELECT * FROM gop_anketcevap ORDER BY answer_id DESC"
set rs = baglanti.execute(SQL)

'set answer id to next number
answ_id = rs("answer_id") + 1

'add answer
SQL = "INSERT INTO gop_anketcevap (poll_id,answer_id,answer) VALUES (" & poll_id & "," & answ_id & ",'" & Request.Form("add_one") & "')"
set rs = baglanti.execute(SQL)

end if

Response.Redirect("anket.asp?sub=edit&id=" & poll_id)

end sub
%>

<%
'delete poll
sub del_answ()
dim poll_id
'get poll id
poll_id = Request.QueryString ("id")

'delete selected answer
SQL = "DELETE FROM gop_anketcevap WHERE answer_id=" & Request.QueryString("answ_id")
set rs = baglanti.execute(SQL)

Response.Redirect("anket.asp?sub=edit&id=" & poll_id)

end sub
%>

<%
'delete poll
sub del()
dim poll_id
'get poll id
poll_id = Request.QueryString ("id")

'delete poll and corresponding answers
SQL = "DELETE FROM gop_anketsoru WHERE id=" & poll_id
set rs = baglanti.execute(SQL)
SQL = "DELETE FROM gop_anketcevap WHERE poll_id=" & poll_id
set rs = baglanti.execute(SQL)

Response.Redirect("anket.asp")

end sub
%>
<!--#include file="alt.asp"-->