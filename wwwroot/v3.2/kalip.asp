<!--#include file="ayar.asp"-->
<!--#include file="db.asp"-->
<!--#include file="inc.asp"-->
<!--#include file="filtre.asp"-->
<!--#include file="baslik.asp"-->
<%
if mdbisim="db/blog.mdb" then
	response.write "<center>"&chr(10)
	response.write "<h1>Veritabanı Yolunu Değiştiriniz.!</h1><br>"&chr(10)
	response.write "<h2>sitenizin veritabanı yolunu güvenlik nedeniyle değiştirmeniz gerekmektedir</h2><br>"&chr(10)
	response.write "Gerekli ayarlar ayar.asp de"
	response.write "<br>Efendy Blog"
	response.write "</center>"
	response.end
end if
%><!--#include file="tema/kalip.asp"-->