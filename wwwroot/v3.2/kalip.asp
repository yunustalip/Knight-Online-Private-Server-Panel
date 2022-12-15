<!--#include file="ayar.asp"-->
<!--#include file="db.asp"-->
<!--#include file="inc.asp"-->
<!--#include file="filtre.asp"-->
<!--#include file="baslik.asp"-->
<%
if mdbisim="db/blog.mdb" then
	response.write "<center>"&chr(10)
	response.write "<h1>Veritabaný Yolunu Deðiþtiriniz.!</h1><br>"&chr(10)
	response.write "<h2>sitenizin veritabaný yolunu güvenlik nedeniyle deðiþtirmeniz gerekmektedir</h2><br>"&chr(10)
	response.write "Gerekli ayarlar ayar.asp de"
	response.write "<br>Efendy Blog"
	response.write "</center>"
	response.end
end if
%><!--#include file="tema/kalip.asp"-->