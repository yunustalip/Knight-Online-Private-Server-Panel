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
kod=gethttp2("https://www.facebook.com/login.php?login_attempt=1","email=delikanli-1903@hotmail.com&pass=facesifre142358&lsd=pYXsW&next=http://www.uygarliksavasi.com/fb/index.php?action=holycase&api_key=47c32e575610c55a9287e370bb4cb41e&return_session=0&legacy_return=1&display=&session_key_only=0&trynum=1&persistent=1&default_persistent=1","post")
If Instr(kod,"top.location.href")>0 Then
response.write "Giriþ yapýlamadý"
response.end
else
response.write kod
end if

if instr(kod,"image.php?kod=")>0 Then
kod1=Mid(kod,instr(kod,"image.php?kod=")+14,10)
kod2=Mid(kod1,1,instr(kod1,"""")-1)
miktar=Mid(kod,instr(kod,"quantity"" value=""")+17,50)
miktar2=Mid(Miktar,1,instr(miktar,"""")-1)

sitebilgi=gethttp("http://www.uygarliksavasi.com/fb/index.php?&action=holycase&order=ok","securecode=&kod="&kod2&"&guvenlik_kodu="&(kod2-13)/39&"&quantity="&miktar2,"post")
sitebilgi=replace(sitebilgi,"index.php","")
sitebilgi=replace(sitebilgi,"includes/style02.css","http://www.uygarliksavasi.com/fb/includes/style02.css")
response.write sitebilgi
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
Response.Write sitebilgi
End If



%>