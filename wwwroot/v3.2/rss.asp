<% with response
.buffer = true
.expires = 0
.contentType = "text/XML"
end with
function RemoveCode(strInput)
      strInput = Replace(strInput,"&", "&amp;", 1, -1, 1)
Removecode = strInput
end function%><?xml version="1.0" encoding="ISO-8859-9" ?> 
<rss version="2.0"><!--#include file="db.asp"--><!--#include file="inc.asp"--><!--#include file="filtre.asp"-->
<channel>
<title><%=sitebaslik%></title>
<link>http://<%=strsite%></link>
<language>tr-TR</language>
<%
rss=Request.QueryString("rss")
if rss="yorumlar" then

set obj2 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from yorum where onay=0 and blog_id<>0 order by tarih DESC"
obj2.open SQL,data,1,3

For p = 1 To 10
if obj2.eof Then exit For

with response
.write("<item>"&chr(13))
.write("<title>"&Left(RemoveCode(Obj2("yorum")),100)&"..</title>"&chr(13))
.write("<link>http://"&strsite&"/"&SEOLink(Obj2("blog_id"))&"</link>"&chr(13))
.write("<pubDate>"&RemoveCode(Obj2("tarih"))&"</pubDate>"&chr(13))
.write("</item>"&chr(13))
end with

Obj2.Movenext
next
Obj2.close
set Obj2 = nothing


elseif rss="blog" then
id=Request.QueryString("id")
	if isnumeric(id)=false then
		response.redirect "index.asp"
	end if
set obj2 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from yorum where onay=0 and blog_id="&id&" order by tarih DESC"
obj2.open SQL,data,1,3

	if obj2.eof then
		response.write "kayıt yok"
	end if

For p = 1 To 10
if obj2.eof Then exit For

with response
.write("<item>"&chr(13))
.write("<title>"&Left(RemoveCode(Obj2("yorum")),100)&"..</title>"&chr(13))
.write("<link>http://"&strsite&"/"&SEOLink(Obj2("blog_id"))&"</link>"&chr(13))
.write("<pubDate>"&RemoveCode(Obj2("tarih"))&"</pubDate>"&chr(13))
.write("</item>"&chr(13))
end with

Obj2.Movenext
next
Obj2.close
set Obj2 = nothing

elseif rss="kategori" then
id=Request.QueryString("id")
	if isnumeric(id)=false then
		response.redirect "index.asp"
	end if
set obj = Server.CreateObject("ADODB.RecordSet")
SQL = "select id,tarih,konu,mesaj from blog where kat_id="&id&" order by id DESC"
obj.open SQL,data,1,3
	if obj.eof then
		response.write "kayıt yok"
	end if
For p = 1 To 10
if obj.eof Then exit For

with response
.write("<item>"&chr(13))
.write("<title>"&RemoveCode(Obj("konu"))&"</title>"&chr(13))
.write("<link>http://"&strsite&"/"&SEOLink(Obj("id"))&"</link>"&chr(13))
.write("<description><![CDATA["&YaziKirp(obj("mesaj"),SEOLink(obj("id")))&"]]></description>")
.write("<pubDate>"&Obj("tarih")&"</pubDate>"&chr(13))
.write("</item>"&chr(13))
end with

obj.Movenext 
Next
obj.Close
Set obj = Nothing


elseif rss="bloglar" then

set obj = Server.CreateObject("ADODB.RecordSet")
SQL = "select id,konu,mesaj,tarih from blog order by id DESC"
obj.open SQL,data,1,3

For p = 1 To 10
if obj.eof Then exit For

with response
.write("<item>"&chr(13))
.write("<title>"&RemoveCode(Obj("konu"))&"</title>"&chr(13))
.write("<link>http://"&strsite&"/"&SEOLink(Obj("id"))&"</link>"&chr(13))
.write("<description><![CDATA["&YaziKirp(obj("mesaj"),SEOLink(obj("id")))&"]]></description>")
.write("<pubDate>"&Obj("tarih")&"</pubDate>"&chr(13))
.write("</item>"&chr(13))
end with

obj.Movenext 
Next
obj.Close
Set obj = Nothing
else
with response
.write("<item>"&chr(13))
.write("<title>Bloglar</title>"&chr(13))
.write("<link>http://"&strsite&"/rss.asp?rss=bloglar</link>"&chr(13))
.write("</item>"&chr(13))
.write("<item>"&chr(13))
.write("<title>Yorumlar</title>"&chr(13))
.write("<link>http://"&strsite&"/rss.asp?rss=yorumlar</link>"&chr(13))
.write("</item>"&chr(13))
end with
end if
data.close
set data = nothing
%></channel></rss>