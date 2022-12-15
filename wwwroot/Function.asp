<%response.charset="iso-8859-9"
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

Function guvenlik(data) 
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
data = Replace (data ,"And","",1,-1,1) 
data = Replace (data ,"'","",1,-1,1) 
data = Replace (data ,"Chr(34)","",1,-1,1) 
data = Replace (data ,"Chr(39)","",1,-1,1) 
data = Replace (data ,"select","",1,-1,1) 
data = Replace (data ,"join","",1,-1,1) 
data = Replace (data ,"union","",1,-1,1) 
data = Replace (data ,"where","",1,-1,1) 
data = Replace (data ,"insert","",1,-1,1) 
data = Replace (data ,"delete","",1,-1,1) 
data = Replace (data ,"Update","",1,-1,1) 
data = Replace (data ,"like","",1,-1,1) 
data = Replace (data ,"drop","",1,-1,1) 
data = Replace (data ,"create","",1,-1,1) 
data = Replace (data ,"modify","",1,-1,1) 
data = Replace (data ,"rename","",1,-1,1) 
data = Replace (data ,"alter","",1,-1,1) 
data = Replace (data ,"cast","",1,-1,1) 
guvenlik=data 
end Function 

function cla(tur)

select case tur
case "101", "105", "106", "201", "205", "206"
cla="Warrior"
case "102", "107", "108", "202", "207", "208"
cla="Rogue"
case "103", "109", "110", "203", "209", "210"
cla="Mage"
case "104", "111", "112", "204", "211", "212"
cla="Priest"
Case else
cla="Unknown"
end select

end function

function cla2(tur2)

select case tur2
case "101", "105", "201", "205"
cla2="Warrior"
case "106"
cla2="Berserker Hero"
case "206"
cla2="Blade Master"
case "102", "107", "202", "207"
cla2="Rogue"
case "108"
cla2="Shadow Vain"
case "208"
cla2="Kasar Hood"
case "103", "109", "203", "209"
cla2="Mage"
case "110"
cla2="Elemental Lord"
case "210"
cla2="Arch Mage"
case "104", "111", "204", "211"
cla2="Priest"
case "112"
cla2="Shadow Knight"
case "212"
cla2="Paladin"
Case else
cla2="Unknown"
end select

end function

function nation(irk)
if irk="1" Then
Response.Write "<img src='imgs/karuslogo.gif' />"
elseif irk="2" Then
Response.Write "<img src='imgs/elmologo.gif' />"
End If

end function

function sirarenk(sira)
if sira="1" Then
style="font-weight:bold; color:red" 
elseif sira="2" Then
style="font-weight:bold; color:black"
elseif sira="3" Then
style="font-weight:bold; color:blue"
else
style="color:black"
End If
sirarenk=style
end function

function resim(resimid)
resimid="itemicon_"+left(resimid,1)+"_"+mid(resimid,2,4)+"_"+mid(resimid,6,2)+"_"+mid(resimid,7,1)+".jpg"
resim=resimid
end function

function resim2(resimid)
if resimid=170250255 or resimid=170210271 or resimid=170210272 or resimid=170210273 or resimid=170210274 or resimid=170210275 or resimid=170210276 or resimid=170210277 or resimid=170210278 or resimid=170210279 or resimid=170210280 Then
resim2="itemicon_1_7155_00_0.jpg"
else
resimid="itemicon_"+left(resimid,1)+"_"+mid(resimid,2,4)+"_"+mid(resimid,6,1)+"0_0.jpg"
resim2=resimid
End If
end function


Function ayir(Veri)
Dim ixy,GeciciVeri
GeciciVeri=""
Veri = Cstr(StrReverse(Veri))
For ixy = 1 To Len(Veri)
GeciciVeri = Mid(Veri,ixy,1) & GeciciVeri
If ixy Mod 3 = 0 And Not ixy = Len(Veri) Then GeciciVeri = "." & GeciciVeri
Next
ayir = GeciciVeri
End Function

Function tr(id)
Dim prc
For prc=1 to len(id)
If instr("abcdefghigklmnoprstuvyzxwq0123456789",mid(id,prc,1))<>0 Then
tr=id
Else
Response.Write "<font color=red><b>Lütfen Özel karakterler Kullanmayýnýz!</b></font><script>var hata = (document.getElementById('usernam').focus()); eval(hata);</script>"
Response.End
Exit Function
Response.End
End If
Next
End function

function vt(gelen)
Conne.Execute(gelen)
end function

function gizlis(gizlisoru)
if gizlisoru="1" Then 
gizlis="Ilk öðretmeninizin ismi nedir?"
elseif gizlisoru="2" Then
gizlis="En çok sevdiðiniz film nedir?"
elseif gizlisoru="3" Then
gizlis="Annenizin kýzlýk soy adý nedir?"
elseif gizlisoru="4" Then
gizlis="Ýlk evcil hayvanýnýzýn ismi nedir?"
elseif gizlisoru="5" Then
gizlis="Çocukken en çok sevdiðiniz yer neresiydi?"
elseif gizlisoru="6" Then
gizlis="Annenizin doðduðu þehir nedir?"
elseif gizlisoru="7" Then
gizlis="En sevdiðiniz kitap nedir?"
elseif gizlisoru="8" Then
gizlis="En sevdiðiniz süper kahraman kimdir?"
End If
end function

Function EmailKontrol(Str)
Et_Isareti = InStr(2, Str , "@" )
If Et_Isareti=0 Then
EmailKontrol=False
Else
Et_Isareti_Krakter_Sayisi=Et_Isareti
Et_Isareti=True
End If
If Et_Isareti = True Then
Nokta = InStr(Et_Isareti_Krakter_Sayisi + 2, Str , "." )
If Nokta=0 Then
EmailKontrol="False"
Else
EmailKontrol="True"
End If
Else
EmailKontrol="False"
End If
End Function

function yn(adres)
Response.Write "<meta http-equiv=""refresh"" content=""1;url="&adres&""">"
Response.End
End Function 


Function QueryFilter(Str)
Str = Replace(Str, "*", "[INJ]",1,-1,1)
Str = Replace(Str, "=", "[INJ]",1,-1,1)
Str = Replace(Str, "<", "[INJ]",1,-1,1)
Str = Replace(Str, ">", "[INJ]",1,-1,1)
Str = Replace(Str, ";", "[INJ]",1,-1,1)
Str = Replace(Str, "(", "[INJ]",1,-1,1)
Str = Replace(Str, ")", "[INJ]",1,-1,1)
Str = Replace(Str, "+", "[INJ]",1,-1,1)
Str = Replace(Str, "#", "[INJ]",1,-1,1)
Str = Replace(Str, "'", "[INJ]", 1, -1, 1)
Str = Replace(Str, "&", "[INJ]", 1, -1, 1)
Str = Replace(Str, "%", "[INJ]", 1, -1, 1)
Str = Replace(Str, "?", "[INJ]", 1, -1, 1)
Str = Replace(Str, "´", "[INJ]", 1, -1, 1)
Str = Replace(Str, ",", "[INJ]",1,-1,1)
Str = Replace(Str, "UNION", "[INJ]",1,-1,1)
Str = Replace(Str, "SELECT", "[INJ]",1,-1,1)
Str = Replace(Str, "WHERE", "[INJ]",1,-1,1)
Str = Replace(Str, "LIKE", "[INJ]",1,-1,1)
Str = Replace(Str, "FROM", "[INJ]",1,-1,1)
Str = Replace(Str, "UPDATE", "[INJ]",1,-1,1)
Str = Replace(Str, "INSERT", "[INJ]",1,-1,1)
Str = Replace(Str, "ORDER", "[INJ]",1,-1,1)
Str = Replace(Str, "GROUP", "[INJ]",1,-1,1)
Str = Replace(Str, "ALTER", "[INJ]",1,-1,1)
Str = Replace(Str, "ADD", "[INJ]",1,-1,1)
Str = Replace(Str, "MODIFY", "[INJ]",1,-1,1)
Str = Replace(Str, "RENAME", "[INJ]",1,-1,1)
Str = Replace(Str, Chr(39), "[INJ]", 1, -1, 1)
If InStr(1,Str,"[INJ]",1) then
Response.Redirect "Default.asp"
end if
QueryFilter = Str
End Function


Private Function emailAddressValidation(strEmailAddress)
	Dim intLoopCounter 	'Holds the loop counter
	
	strEmailAddress = Trim(LCase(strEmailAddress))
	
	strEmailAddress = Replace(strEmailAddress, "..", ".")
	
	For intLoopCounter = 0 to 37
		strEmailAddress = Replace(strEmailAddress, CHR(intLoopCounter), "", 1, -1, 0) 
	Next

	For intLoopCounter = 39 to 42
		strEmailAddress = Replace(strEmailAddress, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	strEmailAddress = Replace(strEmailAddress, CHR(44), "", 1, -1, 0)
	
	For intLoopCounter = 58 to 60
		strEmailAddress = Replace(strEmailAddress, CHR(intLoopCounter), "", 1, -1, 0) 
	Next

	strEmailAddress = Replace(strEmailAddress, CHR(62), "", 1, -1, 0)

	For intLoopCounter = 65 to 94
		strEmailAddress = Replace(strEmailAddress, CHR(intLoopCounter), "", 1, -1, 0) 
	Next

	strEmailAddress = Replace(strEmailAddress, CHR(96), "", 1, -1, 0)

	For intLoopCounter = 123 to 125
		strEmailAddress = Replace(strEmailAddress, CHR(intLoopCounter), "", 1, -1, 0) 
	Next

	For intLoopCounter = 127 to 255
		strEmailAddress = Replace(strEmailAddress, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	If Len(strEmailAddress) < 5 OR NOT Instr(1, strEmailAddress, " ") = 0 OR InStr(1, strEmailAddress, "@", 1) < 2 OR InStrRev(strEmailAddress, ".") < InStr(1, strEmailAddress, "@", 1) Then strEmailAddress = ""
	emailAddressValidation = strEmailAddress

End Function

%>