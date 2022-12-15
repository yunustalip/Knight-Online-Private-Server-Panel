<%
Function Buyut(strBaslik)
strBaslik = Replace(strBaslik, "Ç", "C")
strBaslik = Replace(strBaslik, "Ð", "G")
strBaslik = Replace(strBaslik, "Ý", "I")
strBaslik = Replace(strBaslik, "Ö", "O")
strBaslik = Replace(strBaslik, "Þ", "S")
strBaslik = Replace(strBaslik, "Ü", "U")
Buyut = strBaslik
End Function

Const ServerURL = "http://joyturkaac.radyolarburada.com:9055"
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
		objXmlHttp.sEnd 
		GETHTTP = objXmlHttp.ResponseBody
		Set objXmlHttp = Nothing
	End Function

Verimiz = GETHTTP(ServerURL)
Veriler = BinaryToString(Verimiz)

If Err.Number <> 0 Then
	Gosterim = False
End If

If Instr(Veriler,"Server is currently down") > 0 Then
	Gosterim = False
End If

IF Gosterim = True Then

Basla3="Current Song: </font></td><td><font class=default><b>"
Bitir3="</b></td></tr></table><br><table cellpadding=0 cellspacing=0 border=0 width=100"
Itibaren3 = Instr(Veriler,Basla3)

Sarki = Mid(Veriler,Itibaren3+53,len(Veriler))
Ekadar3 = Instr(Sarki,Bitir3)
CalanSarki = Mid(Sarki,1,Ekadar3-1)


If Instr(CalanSarki,"-") Then
SarkiBilgi=Split(CalanSarki,"-")
Sanatci = SarkiBilgi(0)
Sarki = SarkiBilgi(1)

Resim = "http://www.spectrummedya.com.tr/desktop/"&Trim(Sanatci)&".JPG"


Set XmlHttp = server.CreateObject("MSXML2.ServerXMLHTTP")
XmlHttp.Open "GET", Resim, False
XmlHttp.send
Resim = XmlHttp.ResponseBody
Set XmlHttp = Nothing

Set BinaryStream = Server.CreateObject("ADODB.Stream") 
BinaryStream.Type = 1
BinaryStream.Open 
binarystream.write resim
BinaryStream.SaveToFile server.mappath("\images\"&Trim(Sanatci)&".JPG"), 2
Set BinaryStream = Nothing

Else

End If

End If 

%>