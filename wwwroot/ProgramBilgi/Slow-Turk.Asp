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
End If
%>
<style>
.alpha60 {
	/* Fallback for web browsers that doesn't support RGBa */
	background: rgb(204, 204, 204);
	/* RGBa with 0.6 opacity */
	background: rgba(204, 204, 204, 0.6);
	/* For IE 5.5 - 7*/
	filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#CCCCCC, endColorstr=#CCCCCC);
	/* For IE 8*/
	-ms-filter: "progid:DXImageTransform.Microsoft.gradient(startColorstr=#CCCCCC, endColorstr=#CCCCCC)";
}
</style>

<div style="float:left"><strong>Frekanslar: | </strong></div> <marquee behavior="scroll" direction="left" scrollamount="2" width="248"><strong> İstanbul </strong>95.3 - <strong> Ankara</strong> 95.3 - <strong> İzmir</strong> 99.7 - <strong>Adana</strong> 98.5 - <strong>Antalya </strong> 93.2 - <strong>Bursa</strong> 88.1 - <strong>Çorlu</strong> 95.3 - <strong>Eskişehir</strong> 93.0 -<strong>Gaziantep</strong> 107.3 - <strong>Mersin</strong> 87.9 - <strong>Konya</strong> 94.5 - <strong>Yalova</strong> 95.3 -<strong> Balıkesir</strong> 88.0 -<strong> Kocaeli</strong> 105.5 -<strong> Samsun</strong> 105.7</marquee>
<br><div align="center"><a href="http://www.slowturk.com.tr/yayinakisi.asp" target="_blank"><div style="color:#333333; text-decoration:overline underline">Yayın Akışı</div></a></div>
<h3>Top 20</h3></center>
<iframe src="http://www.slowturk.com.tr/top20_ic.asp" height="75" width="275"></iframe>
</center>
<div class="alpha60" align="center"> <h3><strong><em><u>İstek Hattı</u></em></strong></h6>
İstekleriniz Hafta içi hergün 11:00 - 15:00 saatleri arasında  <strong>''Aşkın Sesi Cenk'' </strong> ile sevdiklerinize ulaşıyor.<br/><br/>
Cep telefonunuzun mesaj bölümünün ilk kısmına <strong>slow yazın, 3004'e yollayın. </strong><br />
  Ya da <strong>cenk@slowturkradyo.com</strong> adresine mail atın. 
</div>