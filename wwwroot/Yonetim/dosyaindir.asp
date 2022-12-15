<%if Session("durum")="esp" Then 
response.charset="iso-8859-9"%><html >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" />
<title>Asp ile dosya indir</title>
<style type="text/css">
body {
	font-family: arial;
	font-size: 12px;
	color: #333;
}
.input {
	width: 200px;
	border:1px solid #CCC;
	margin: 5px 0 0 0;
}
.buton {
	width: 108px;
	border:1px solid #CCC;
	margin: 5px 0 0 0;
}
.select {
	border:1px solid #CCC;
	margin: 0;
}
</style>
</head>
<body>
<form action="?indir=lan" method="post">
Dosya url: <input class="input" type="text" name="url" />
<button class="buton" type="submit">İndir</button>
</form>
<%
indir = Request.Querystring("indir")
if indir = "lan" Then

dosya_adresi    = request.form("url")

if dosya_adresi = ""  Then
	Response.Write "Boş alan bırakmayınız."
    Response.End
End If

Set XmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
XmlHttp.Open "GET", dosya_adresi, False
XmlHttp.send
Peksoy = XmlHttp.ResponseBody
Set XmlHttp = Nothing

Set BinaryStream = Server.CreateObject("ADODB.Stream")
BinaryStream.Type = 1
BinaryStream.Open
BinaryStream.Write Peksoy
BinaryStream.SaveToFile Server.MapPath("../Uploads/"&right(dosya_adresi,(len(dosya_adresi)-instrrev(dosya_adresi,"/")))), 2
Set BinaryStream = Nothing

Response.Write "<strong>"& right(dosya_adresi,(len(dosya_adresi)-instrrev(dosya_adresi,"/"))) &"</strong> adındaki dosya <strong>Uploads</strong> klasörüne indirildi."

End If

else
Response.Redirect("default.asp")
End If%>
</body>
</html>
