<%Session.CodePage=65001
response.charset="utf-8"

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

Const ServerURL = "http://www.powerturk.com/2010/m1.asp"
Gosterim = True
Private  Function BinaryToString(Binary)
Dim  cl1, cl2, cl3, pl1, pl2, pl3
Dim  L
cl1 = 1
cl2 = 1
cl3 = 1
L = LenB(Binary)
Do  While cl1<=L
pl3 = pl3 &  Chr(AscB(MidB(Binary,cl1,1)))
cl1 = cl1 + 1
cl3 = cl3 + 1
If  cl3>300  Then
pl2 = pl2 & pl3
pl3 = ""
cl3 = 1
cl2 = cl2 + 1
If  cl2>200  Then
pl1 = pl1 & pl2
pl2 = ""
cl2 = 1
End  If
End  If
Loop
BinaryToString = pl1 & pl2 & pl3
End  Function
Err.Clear
On Error Resume Next
	Function GETHTTP(adres) 
		Set objXmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objXmlHttp.Open "GET" , adres , false
		objXmlHttp.setRequestHeader "Content-Type", "text/HTML"
		objXmlHttp.setRequestHeader "CharSet", "UTF-8"
		objXmlHttp.sEnd 
		GETHTTP = objXmlHttp.Responsetext
		Set objXmlHttp = Nothing
	End Function

Verimiz = GETHTTP(ServerURL)
Veriler = BinaryToString(Verimiz)

If Err.Number <> 0 Then
	Gosterim = False
End If

If Instr(Veriler,"asd") > 0 Then
	Gosterim = False
End If

	IF Gosterim = True Then

Basla3="<font face=""arial"" size=""1"" color=""000000"">&nbsp;&nbsp;<b>"
Bitir3="</font>"
Itibaren3 = Instr(Veriler,Basla3)

Sarki = Mid(Veriler,Itibaren3+len(basla3),len(Veriler))
Ekadar3 = Instr(Sarki,Bitir3)
CalanSarki = Mid(Sarki,1,Ekadar3-1)
%><table  style="font-size:10px;color:#333333" cellspacing="0" cellpadding="0">
<tr><td>
<strong>Telefon: </strong>0 216 554 0 404 <br>
<strong>Bilgi için: </strong>hello@powerturk.com <br>
<strong>Teknik sorunlar için: </strong>powerteknik@powerturk.com 
</td>
</tr>
<tr><td><center><b>Program:
<%Response.Write "<font face=""verdana"" size=""1"">"& CalanSarki &"" %></b>
</center></td></tr>
</table>
<%
End If
%>