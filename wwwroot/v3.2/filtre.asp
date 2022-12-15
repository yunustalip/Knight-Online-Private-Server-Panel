<%
Function AramaFiltre(efendy)
	efendy=Replace(efendy,"%3C","&lt;")
	efendy=Replace(efendy,"%3E","&gt;")
	efendy=Replace(efendy,"'","&#39;")
	efendy=Replace(efendy,"%22","&quot;")
	efendy=Replace(efendy,"""","&#34;")
	AramaFiltre=efendy
End Function

Function  filtre(Yasak)
if Yasak="" Then Exit Function
     	Yasak = Replace(Yasak, "<", "&lt;")
     	Yasak = Replace(Yasak, ">", "&gt;")
     	Yasak = Replace(Yasak, "=", "&#061;", 1, -1, 1)
     	Yasak = Replace(Yasak, "'", "&#39;", 1, -1, 1)
     	Yasak = Replace(Yasak, "select", "sel&#101;ct", 1, -1, 1)
     	Yasak = Replace(Yasak, "join", "jo&#105;n", 1, -1, 1)
     	Yasak = Replace(Yasak, "union", "un&#105;on", 1, -1, 1)
     	Yasak = Replace(Yasak, "where", "wh&#101;re", 1, -1, 1)
     	Yasak = Replace(Yasak, "insert", "ins&#101;rt", 1, -1, 1)
     	Yasak = Replace(Yasak, "delete", "del&#101;te", 1, -1, 1)
     	Yasak = Replace(Yasak, "update", "up&#100;ate", 1, -1, 1)
     	Yasak = Replace(Yasak, "like", "lik&#101;", 1, -1, 1)
     	Yasak = Replace(Yasak, "drop", "dro&#112;", 1, -1, 1)
     	Yasak = Replace(Yasak, "create", "cr&#101;ate", 1, -1, 1)
     	Yasak = Replace(Yasak, "modify", "mod&#105;fy", 1, -1, 1)
     	Yasak = Replace(Yasak, "rename", "ren&#097;me", 1, -1, 1)
     	Yasak = Replace(Yasak, "alter", "alt&#101;r", 1, -1, 1)
     	Yasak = Replace(Yasak, "cast", "ca&#115;t", 1, -1, 1)
filtre = Yasak
End  Function

Function MesajFormatla(strMesaj)
     strMesaj = Replace(strMesaj, ":)", "<img src=""tema/"&tema&"/images/smileys/smile.gif"">", 1, -1, 1) 
     strMesaj = Replace(strMesaj, ":(", "<img src=""tema/"&tema&"/images/smileys/frown.gif"">", 1, -1, 1) 
     strMesaj = Replace(strMesaj, ":D", "<img src=""tema/"&tema&"/images/smileys/biggrin.gif"">", 1, -1, 1) 
     strMesaj = Replace(strMesaj, ":o:", "<img src=""tema/"&tema&"/images/smileys/redface.gif"">", 1, -1, 1) 
     strMesaj = Replace(strMesaj, ";)", "<img src=""tema/"&tema&"/images/smileys/wink.gif"">", 1, -1, 1) 
     strMesaj = Replace(strMesaj, ":p", "<img src=""tema/"&tema&"/images/smileys/tongue.gif"">", 1, -1, 1) 
     strMesaj = Replace(strMesaj, ":cool:", "<img src=""tema/"&tema&"/images/smileys/cool.gif"">", 1, -1, 1) 
     strMesaj = Replace(strMesaj, ":rolleyes:", "<img src=""tema/"&tema&"/images/smileys/rolleyes.gif"">", 1, -1, 1) 
     strMesaj = Replace(strMesaj, ":mad:", "<img src=""tema/"&tema&"/images/smileys/mad.gif"">", 1, -1, 1) 
     strMesaj = Replace(strMesaj, ":eek:", "<img src=""tema/"&tema&"/images/smileys/eek.gif"">", 1, -1, 1) 
     strMesaj = Replace(strMesaj, ":confused:", "<img src=""tema/"&tema&"/images/smileys/confused.gif"">", 1, -1, 1) 
     strMesaj = Replace(strMesaj, vbCrlf, vbCrlf&"<br />", 1, -1, 1)
     MesajFormatla = strMesaj
End  Function

Function CreateURL(ByVal strVariable)

	Dim strTempURL

	strTempURL = Trim(strVariable)

	'// Türkçe karakterler deðiþtiriliyor
	strTempURL = Replace(strTempURL,"ç","c")
	strTempURL = Replace(strTempURL,"Ç","c")
	strTempURL = Replace(strTempURL,"ð","g")
	strTempURL = Replace(strTempURL,"Ð","g")
	strTempURL = Replace(strTempURL,"Ý","i")
	strTempURL = Replace(strTempURL,"I","i")
	strTempURL = Replace(strTempURL,"ý","i")
	strTempURL = Replace(strTempURL,"ö","o")
	strTempURL = Replace(strTempURL,"Ö","o")
	strTempURL = Replace(strTempURL,"þ","s")
	strTempURL = Replace(strTempURL,"Þ","s")
	strTempURL = Replace(strTempURL,"ü","u")
	strTempURL = Replace(strTempURL,"Ü","u")

	strTempURL = CleanChars(strTempURL)
	strTempURL = LCase(Replace(strTempURL, " ", "-", 1, -1, 1))

	'// -- karakteri temizleniyor
	Do While InStr(strTempURL, "--")
		strTempURL = Replace(strTempURL, "--", "-", 1, -1, 1)
	Loop

	CreateURL = Left(strTempURL,50)

End Function

Function CleanChars(ByVal strVariable)

	Dim objRegExp
	Dim strTempValue

	Set objRegExp = New RegExp
	With objRegExp
		.Pattern = "([^a-zA-Z0-9])"
		.IgnoreCase = False
		.Global = True
		strTempValue = .Replace(strVariable, " ")
	End With

	CleanChars = strTempValue

End Function

Function SEOLink(efendy)
	if not isnumeric(efendy)=false then
		if Strseoayar=2 then
			efendy="blog.asp?id="&efendy
		else
			set rs = Server.CreateObject("ADODB.RecordSet")
			SQL = "select konu from blog where id="&efendy
			rs.open SQL,data,1,3
			if not rs.eof then
			efendy=efendy&"-"&CreateURL(rs("konu"))&".html"
			end if
			rs.close : set rs=nothing
		end if
	end if
SEOLink=efendy
End Function

Function Cevir(strVeri) 

     If strVeri = "" Then Exit Function 

     Set objRegExp = New Regexp 
     With objRegExp 
          .Pattern = "<.*?>" 
          .IgnoreCase = False 
          .Global = True 
     End With 

     Cevir = objRegExp.Replace(strVeri,"") 

End Function 

Function YaziKirp(efendy,Elink)
	if Stryazikes="1" then
		if not InStr(efendy,"{KES}")>0 then
		if isnumeric(Stryaziuzunluk)=false then : uzunluk="500" : else : uzunluk=Stryaziuzunluk : end if
		efendy=Left(Cevir(efendy),uzunluk)&"<a href="&Elink&">..devamý&gt;&gt;</a>"
		else
		bitir=instr(efendy,"{KES}")-1
		efendy=Mid(efendy,1,bitir)&"<a href="&Elink&">..devamý&gt;&gt;</a>"
		end if
	else
	if InStr(efendy,"{KES}")>0 then
		bitir=instr(efendy,"{KES}")-1
		efendy=Mid(efendy,1,bitir)&"<a href="&Elink&">..devamý&gt;&gt;</a>"
	end if
	end if

	efendy=Replace(efendy,"{KES}","")
YaziKirp=efendy
End Function
%>