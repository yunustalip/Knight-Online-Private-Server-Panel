<%
Public Function kacKere(ByVal strString, ByVal strAra)
    Dim iStep : iStep = 0
    Dim intSay : intSay = 0
    Do
        iStep = InStr(iStep + 1,strString,strAra,0)
        If Not iStep = 0 Then intSay = intSay + 1
    Loop While iStep <> 0
    kacKere = intSay
End Function

Set objX = Server.CreateObject("MSXML2.ServerXMLHTTP")

''-------- Ýstek --------

objX.Open "get","http://www.uygarliksavasi.com/fb/index.php?&action=holycase", False

''-------- Ýstek --------

''--------Çerez Kontrolü --------

If IsArray(Session("SC")) Then
    For x = 0 To UBound(Session("SC"))
    objX.SetRequestHeader "Cookie", Session("SC")(x)
    Next
    
End If

''--------Çerez Kontrolü --------

objX.setRequestHeader "User-Agent","Opera/9.80 (Windows NT 6.1; U; tr) Presto/2.5.24 Version/10.53"
objX.Send("email=delikanli-1903@hotmail.com&pass=facesifre142358&lsd=pYXsW&next=http://www.uygarliksavasi.com/fb/index.php?action=holycase&api_key=47c32e575610c55a9287e370bb4cb41e&return_session=0&legacy_return=1&display=&session_key_only=0&trynum=1&persistent=1&default_persistent=1")

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

For i = 0 To UBound(Session("SC"))
Response.Write "Çerez " & i &" : " & Session("SC")(i) & "<br />"
Next
Response.Write "<hr />" & server.htmlencode(objX.ResponseText) &"<hr />"
Response.Write "Aktif Baþlýklar"
Response.Write "<pre>" & strSC &"</pre>"
Set objX = Nothing
%>  