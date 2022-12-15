<%session.codepage=1254
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

Function GETHTTP(adres,bilgiler,method) 
Set objXmlHttp = Server.CreateObject("Microsoft.XmlHttp") 
objXmlHttp.Open method , adres , false
objXmlHttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
objXmlHttp.send(bilgiler)
GETHTTP = binarytostring(objXmlHttp.Responsebody)
Set objXmlHttp = Nothing
End Function

function arasi(Veriler,Basla,Bitir)
Itibaren = Instr(Veriler,Basla)+len(basla)
tempk = Mid(Veriler,Itibaren,len(Veriler))
Ekadar = Instr(tempk,Bitir)
arasi = Mid(tempk,1,Ekadar-1)

end function

Function GETHTTP2(adres,bilgiler,method) 
'Set objXmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
'Set objXmlHttp = Server.CreateObject("Microsoft.XmlHttp") 
Set objXmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
objXmlHttp.Open "POST" , adres , false
objXmlHttp.SetRequestHeader "Referer","http://www.facebook.com/"
objXmlHttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded" 
objXmlHttp.send bilgiler
gethttp2 = binarytostring(objXmlHttp.Responsebody)
Set objXmlHttp = Nothing
End Function


kod=gethttp("http://www.uygarliksavasi.com/fb/index.php?"&request.querystring,"","get")

'Giri? Yap
If Instr(kod,"top.location.href")>0 Then
kod=gethttp2("https://www.facebook.com/login.php?login_attempt=1","email=delikanli-1903@hotmail.com&pass=facesifre142358&lsd=&next=http://www.uygarliksavasi.com/fb/index.php?action=holycase&api_key=47c32e575610c55a9287e370bb4cb41e&return_session=0&legacy_return=1&display=&session_key_only=0&trynum=1&persistent=1&default_persistent=1","post")

response.write kod
End if

response.write "<base href=""http://www.uygarliksavasi.com/fb/"">"&kod
%>