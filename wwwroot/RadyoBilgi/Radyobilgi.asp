<%response.charset="utf-8"
function secur(data) 
Data = Replace( data , "'" , "", 1, -1,1)
data = Replace (data ,"`","",1,-1,1) 
data = Replace (data ,"=","",1,-1,1) 
data = Replace (data ,"&","",1,-1,1) 
data = Replace (data ,"%","",1,-1,1) 
data = Replace (data ,"!","",1,-1,1) 
data = Replace (data ,"#","",1,-1,1) 
data = Replace (data ,"<","",1,-1,1) 
data = Replace (data ,">","",1,-1,1) 
data = Replace (data ,"*","",1,-1,1) 
data = Replace (data ,"'","",1,-1,1) 
data = Replace (data ,"Chr(34)","",1,-1,1)
data = Replace (data ,"Chr(39)","",1,-1,1)
secur=data 
end function
ServerURL = secur(Request.Querystring("url"))

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
End If

IF Gosterim = True Then
Response.Write CalanSarki
End If
%>