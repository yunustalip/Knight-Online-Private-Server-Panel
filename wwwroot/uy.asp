<%session.codepage=1254
on error resume next

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
' Set objXmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
' Set objXmlHttp = Server.CreateObject("Microsoft.XmlHttp")
Set objXmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
objXmlHttp.Open method , adres , false
objXmlHttp.setRequestHeader "Referer","http://www.facebook.com/"
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

Public Function kacKere(ByVal strString, ByVal strAra)
    Dim iStep : iStep = 0
    Dim intSay : intSay = 0
    Do
        iStep = InStr(iStep + 1,strString,strAra,0)
        If Not iStep = 0 Then intSay = intSay + 1
    Loop While iStep <> 0
    kacKere = intSay
End Function


Function GETHTTP2(adres,bilgiler,method) 


Set objX = Server.CreateObject("MSXML2.ServerXMLHTTP")

''-------- İstek --------

objX.Open "POST" , adres , false

''-------- İstek --------

''--------Çerez Kontrolü --------

If IsArray(Session("SC")) Then
    For x = 0 To UBound(Session("SC"))
    objX.SetRequestHeader "Cookie", Session("SC")(x)
    Next
    
End If

''--------Çerez Kontrolü --------
objX.SetRequestHeader "Referer","http://www.facebook.com/"
objX.setRequestHeader "Content-Type","application/x-www-form-urlencoded" 
objX.setRequestHeader "User-Agent","Opera/9.80 (Windows NT 6.1; U; tr) Presto/2.5.24 Version/10.53"
objX.Send bilgiler

If Session("SC_All") <> objX.getAllResponseHeaders() Then
Session("SC_All") = objX.getAllResponseHeaders()
End If

Dim strSC : strSC = objX.getAllResponseHeaders()

''-------- Oturumda Çerez Saklama --------

If Not IsArray(Session("SC")) Then 
    Dim arrSC()
    
    ReDim arrSC(kacKere(strSC,"Set-Cookie:") - 1)
    Dim iSay : iSay = 0
    splSC = Split(strSC,vbCrLf)
    For b = 0 To UBound(splSC)
        If Left(Trim(splSC(b)),11) = "Set-Cookie:" Then
            arrSC(iSay) = Replace(splSC(b),"Set-Cookie:","")
            arrSC(iSay) = Trim(arrSC(iSay))
            iSay = iSay + 1
        End If
    Next
    Session("SC") = arrSC
End If

''-------- Oturumda Çerez Saklama --------


gethttp2 = binarytostring(objX.Responsebody)

End Function


Function fnSort(aSort,arr,arrr,arrrr,link)
Dim intTempStore1,intTempStore2,intTempStore3,intTempStore4,intTempStore5
Dim i, j 
For i = 1 To UBound(aSort) - 1
For j = i To UBound(aSort)

If aSort(i) < aSort(j) Then

intTempStore = aSort(i)
intTempStore2 = arr(i)
intTempStore3 = arrr(i)
intTempStore4 = arrrr(i)
intTempStore5 = link(i)
aSort(i) = aSort(j)
aSort(j) = intTempStore
arr(i) = arr(j)
arr(j) = intTempStore2
arrr(i) = arrr(j)
arrr(j) = intTempStore3
arrrr(i) = arrrr(j)
arrrr(j) = intTempStore4
link(i) = link(j)
link(j) = intTempStore5
End If 

Next 
Next 
fnSort = aSort
End Function 

kod=gethttp("http://www.uygarliksavasi.com/fb/index.php?&action=holycase","","get")

'Giris Yap
If Instr(kod,"top.location.href")>0 Then
kod=gethttp2("http://www.facebook.com/login.php?login_attempt=1","email=delikanli-1903@hotmail.com&pass=facesifre142358&lsd=&post_form_id=ae15b2e2cac8401714f442f505c1d1da&charset_test=€,´,€,´,水,Д,Є&version=1&ajax=0&width=0&pxr=0&gps=0&m_ts=1329514237&li=_cY-T7eHKQQ3qwn7PwkAA3hO&login=","post")

response.write (kod)

if instr(kod,"image.php?kod=")>0 Then
kod1=Mid(kod,instr(kod,"image.php?kod=")+14,10)
kod2=Mid(kod1,1,instr(kod1,"""")-1)
miktar=Mid(kod,instr(kod,"quantity"" value=""")+17,50)
miktar2=Mid(Miktar,1,instr(miktar,"""")-1)

sitebilgi=gethttp("http://www.uygarliksavasi.com/fb/index.php?&action=holycase&order=ok","securecode=&kod="&kod2&"&guvenlik_kodu="&(kod2-13)/39&"&quantity="&miktar2,"post")
sitebilgi=replace(sitebilgi,"index.php","")
sitebilgi=replace(sitebilgi,"includes/style02.css","http://www.uygarliksavasi.com/fb/includes/style02.css")
response.write server.htmlencode(sitebilgi)
End If
Response.End
End If
'--------

if instr(kod,"image.php?kod=")>0 Then
kod1=Mid(kod,instr(kod,"image.php?kod=")+14,10)
kod2=Mid(kod1,1,instr(kod1,"""")-1)
miktar=Mid(kod,instr(kod,"quantity"" value=""")+17,50)
miktar2=Mid(Miktar,1,instr(miktar,"""")-1)

sitebilgi=Replace(gethttp("http://www.uygarliksavasi.com/fb/index.php?&action=holycase&order=ok","securecode=&kod="&kod2&"&guvenlik_kodu="&(kod2-13)/39&"&quantity="&miktar2,"post"),"includes/style02.css","http://www.uygarliksavasi.com/fb/includes/style02.css")
sitebilgi=replace(sitebilgi,"index.php","")
Response.Write server.htmlencode(sitebilgi)
End If



%>