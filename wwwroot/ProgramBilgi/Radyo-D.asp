<%response.charset="utf-8"

REFERER_URL = Request.ServerVariables("HTTP_REFERER")
Dosyaismi = lcase(Request.ServerVariables("Script_Name"))

If InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
Else
REFERER_DOMAIN = Left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If


If REFERER_DOMAIN="http://fmradyodinle.net" or  REFERER_DOMAIN="http://www.fmradyodinle.net" or dosyaismi="/default.asp" or dosyaismi="/404.asp"  Then
Else
Response.Clear
Response.Write "<a href=""http://www.FmRadyoDinle.net"">www.FmRadyoDinle.net</a>"
Response.End
End If

If Instr(Request.ServerVariables("ALL_HTTP"),"HTTP_X_REQUESTED_WITH:")>0  or dosyaismi="/default.asp" or dosyaismi="/404.asp" Then
Else
Response.Clear
Response.Write "<a href=""http://www.FmRadyoDinle.net"">www.FmRadyoDinle.net</a>"
Response.End
End If%>
<style>
.tepe {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #005CA9;
	text-decoration: none;
	font-weight: bold;
}

</style>
<font class="tepe">> Tel: 0 212 413 53 39 > E-mail: info@radyod.com <br>&gt; Frekanslar <iframe src="http://www.radyod.com/frekans_ic.htm" frameborder="0" width="100" height="15" scrolling="no" align="absmiddle"></iframe><br>> <a href="http://www.radyod.com/yayinakisi.asp" target="_blank" class="tepe">Yayın Akışı</a><br>
&gt; <a href="http://www.radyod.com/videolar.asp" target="_blank" class="tepe">Videolar</a><br>> <a href="http://www.radyod.com/yenialbum.asp" class="tepe" target="_blank">Ayın Albümü</a><br>> <a href="http://www.radyod.com/yeni.asp" class="tepe" target="_blank">Yeni Çıkanlar</a><br>> Top 40<br><iframe name="bolum1" id="bolum1" src="http://www.radyod.com/top40.asp" frameborder="0" width="254" height="100" scrolling="yes" ></iframe><br><iframe name="uye" id="uye" src="http://www.radyod.com/uye.htm" frameborder="0" width="261" height="30" scrolling="no" ></iframe></font>