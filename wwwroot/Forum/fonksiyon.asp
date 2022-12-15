


<%
'*******************************************************
' Kodlarýmý kullandýðýnýz için teþekkürler
' Kullandýðýnýz siteyi bana bildirirseniz sevinirim
' Efkan 
' email :info@aywebhizmetleri.com
' web sayfalarýmý ziyaret etmeyi unutmayýnýz  
' http://www.makineteknik.com
' http://www.binbirkonu.com
' http://www.aywebhizmetleri.com
' http://www.tekrehberim.com
' http://www.hitlinkler.com
' Size uygun bir web sitem mutlaka vardýr ...
' LÜTFEN BU TÜR ÇALIÞMALARIN ÖNÜNÜ KESMEMEK ÝÇÝN TELÝF YAZILARINI SÝLMEYÝN
' EMEÐE SAYGI LÜTFEN 
' KÝÞÝSEL KULLANIM ÝÇÝN ÜCRETSÝZDÝR DÝÐER KULLANIMLARDA HAK TALEP EDÝLEBÝLÝR
'*******************************************************
%>








<%
function hataver(hata)  ' KULLANIMI hataver("Email Adresi veya Þifrenizi yazmadýnýz.")
hataver="<P><IMG SRC=""images/alert.gif"" WIDTH=""32"" HEIGHT=""32"" BORDER=""0"" ><P>"
hataver=hataver & "<B>" & hata & "</B>"
hataver=hataver & "<P><a href=""javascript:history.back(1)"">"
hataver=hataver & "<IMG SRC=""images/undo.gif"" WIDTH=""20"" BORDER=""0"" ></a>"
hataver=hataver & "&nbsp;&nbsp;<A HREF=""default.asp"">"
hataver=hataver & "<IMG SRC=""images/home.gif"" WIDTH=""20"" BORDER=""0"" ></a>"
Response.Write hataver
End  Function 

function bilgiver(bilgi)  ' KULLANIMI bilgiver("Email Adresi veya Þifrenizi yazmadýnýz.")
bilgiver="<P><IMG SRC=""images/info.gif"" WIDTH=""20"" HEIGHT=""20"" BORDER=""0"" ><P>"
bilgiver=bilgiver & "<B>" & bilgi & "</B>"
bilgiver=bilgiver & "<P><A HREF=""default.asp"">"
bilgiver=bilgiver & "<IMG SRC=""images/home.gif"" WIDTH=""20"" BORDER=""0"" ></a>"
Response.Write bilgiver
End  function

function editor(strForm)
strForm = Replace(strForm, Chr(13), "<br>")
strForm = replace(strForm, "[p]", "<p>")
strForm = replace(strForm, "[b]", "<b>")
strForm = replace(strForm, "[/b]", "</b>")
strForm = replace(strForm, "[i]", "<i>")
strForm = replace(strForm, "[/i]", "</i>")
strForm = replace(strForm, "[center]", "<center>")
strForm = replace(strForm, "[/center]", "</center>")
strForm = replace(strForm, "[hr]", "<hr>")
strForm = replace(strForm, "[u]", "<u>")
strForm = replace(strForm, "[/u]", "</u>")
strForm = replace(strForm, "[1]", "<img src=editor/s1.gif>")
strForm = replace(strForm, "[2]", "<img src=editor/s2.gif>")
strForm = replace(strForm, "[3]", "<img src=editor/s3.gif>")
strForm = replace(strForm, "[link#", "<a target='_blank' href=' ")
strForm = replace(strForm, "[/link]", "</a>")
strForm = replace(strForm, "[img#", "<img border='0' src='")
strForm = replace(strForm, "[email#", "<a href='mailto:")
strForm = replace(strForm, "[/email]", "</a>")
strForm = replace(strForm, "[list]", "<li>")
strForm = replace(strForm, "[/list]", "</li>")
strForm = replace(strForm, "#]", "'>")
editor = strForm
end function


function terseditor(strForm)
strForm = replace(strForm, "<p>", "[p]")
strForm = Replace(strForm, "<br>", Chr(13))
strForm = replace(strForm, "<b>", "[b]")
strForm = replace(strForm, "</b>", "[/b]")
strForm = replace(strForm, "<i>", "[i]")
strForm = replace(strForm, "</i>", "[/i]")
strForm = replace(strForm, "<center>", "[center]")
strForm = replace(strForm, "</center>", "[/center]")
strForm = replace(strForm, "<hr>", "[hr]")
strForm = replace(strForm, "<u>", "[u]")
strForm = replace(strForm, "</u>", "[/u]")
strForm = replace(strForm, "<img src=editor/s1.gif>", "[1]")
strForm = replace(strForm, "<img src=editor/s2.gif>", "[2]")
strForm = replace(strForm, "<img src=editor/s3.gif>", "[2]")
strForm = replace(strForm, "<a target='_blank' href=' ", "[link#")
strForm = replace(strForm, "</a>", "[/link]")
strForm = replace(strForm, "<img border='0' src='", "[img#")
strForm = replace(strForm, "<a href='mailto:", "[email#")
strForm = replace(strForm, "</a>", "[/email]")
strForm = replace(strForm, "<li>", "[list]")
strForm = replace(strForm, "</li>", "[/list]")
strForm = replace(strForm, "'>", "#]")
terseditor = strForm
end function




Private Function temizle(ByVal data)
    data = Replace(data, "</script>", "", 1, -1, 1)
	data = Replace(data, "<script language=""javascript"">", "", 1, -1, 1)
	data = Replace(data, "<script language=javascript>", "", 1, -1, 1)
	data = Replace(data, "script", "&#115;cript", 1, -1, 0)
	data = Replace(data, "SCRIPT", "&#083;CRIPT", 1, -1, 0)
	data = Replace(data, "Script", "&#083;cript", 1, -1, 0)
	data = Replace(data, "script", "&#083;cript", 1, -1, 1)
	data = Replace(data, "object", "&#111;bject", 1, -1, 0)
	data = Replace(data, "OBJECT", "&#079;BJECT", 1, -1, 0)
	data = Replace(data, "Object", "&#079;bject", 1, -1, 0)
	data = Replace(data, "object", "&#079;bject", 1, -1, 1)
	data = Replace(data, "applet", "&#097;pplet", 1, -1, 0)
	data = Replace(data, "APPLET", "&#065;PPLET", 1, -1, 0)
	data = Replace(data, "Applet", "&#065;pplet", 1, -1, 0)
	data = Replace(data, "applet", "&#065;pplet", 1, -1, 1)
	data = Replace(data, "embed", "&#101;mbed", 1, -1, 0)
	data = Replace(data, "EMBED", "&#069;MBED", 1, -1, 0)
	data = Replace(data, "Embed", "&#069;mbed", 1, -1, 0)
	data = Replace(data, "embed", "&#069;mbed", 1, -1, 1)
	data = Replace(data, "event", "&#101;vent", 1, -1, 0)
	data = Replace(data, "EVENT", "&#069;VENT", 1, -1, 0)
	data = Replace(data, "Event", "&#069;vent", 1, -1, 0)
	data = Replace(data, "event", "&#069;vent", 1, -1, 1)
	data = Replace(data, "document", "&#100;ocument", 1, -1, 0)
	data = Replace(data, "DOCUMENT", "&#068;OCUMENT", 1, -1, 0)
	data = Replace(data, "Document", "&#068;ocument", 1, -1, 0)
	data = Replace(data, "document", "&#068;ocument", 1, -1, 1)
	data = Replace(data, "cookie", "&#099;ookie", 1, -1, 0)
	data = Replace(data, "COOKIE", "&#067;OOKIE", 1, -1, 0)
	data = Replace(data, "Cookie", "&#067;ookie", 1, -1, 0)
	data = Replace(data, "cookie", "&#067;ookie", 1, -1, 1)
	data = Replace(data, "form", "&#102;orm", 1, -1, 0)
	data = Replace(data, "FORM", "&#070;ORM", 1, -1, 0)
	data = Replace(data, "Form", "&#070;orm", 1, -1, 0)
	data = Replace(data, "form", "&#070;orm", 1, -1, 1)
	data = Replace(data, "iframe", "i&#102;rame", 1, -1, 0)
	data = Replace(data, "IFRAME", "I&#070;RAME", 1, -1, 0)
	data = Replace(data, "Iframe", "I&#102;rame", 1, -1, 0)
	data = Replace(data, "iframe", "i&#102;rame", 1, -1, 1)
	data = Replace(data, "textarea", "&#116;extarea", 1, -1, 0)
	data = Replace(data, "TEXTAREA", "&#84;EXTAREA", 1, -1, 0)
	data = Replace(data, "Textarea", "&#84;extarea", 1, -1, 0)
	data = Replace(data, "textarea", "&#84;extarea", 1, -1, 1)
	data = Replace(data, "on", "&#111;n", 1, -1, 0)
	data = Replace(data, "ON", "&#079;N", 1, -1, 0)
	data = Replace(data, "On", "&#079;n", 1, -1, 0)
	data = Replace(data, "on", "&#111;n", 1, -1, 1)

	data = Replace(data, "<STR&#079;NG>", "<strong>", 1, -1, 1)
	data = Replace(data, "<str&#111;ng>", "<strong>", 1, -1, 1)
	data = Replace(data, "</STR&#079;NG>", "</strong>", 1, -1, 1)
	data = Replace(data, "</str&#111;ng>", "</strong>", 1, -1, 1)
	data = Replace(data, "f&#111;nt", "font", 1, -1, 0)
	data = Replace(data, "F&#079;NT", "FONT", 1, -1, 0)
	data = Replace(data, "F&#111;nt", "Font", 1, -1, 0)
	data = Replace(data, "f&#079;nt", "font", 1, -1, 1)
	data = Replace(data, "f&#111;nt", "font", 1, -1, 1)
	data = Replace(data, "m&#111;no", "mono", 1, -1, 0)
	data = Replace(data, "M&#079;NO", "MONO", 1, -1, 0)
	data = Replace(data, "M&#079;no", "Mono", 1, -1, 0)
	data = Replace(data, "m&#079;no", "mono", 1, -1, 1)
	data = Replace(data, "m&#111;no", "mono", 1, -1, 1)

	data = Replace(data, "<", "&lt;")
	data = Replace(data, ">", "&gt;")
	'data = Replace(data, "[", "&#091;")
	'data = Replace(data, "]", "&#093;")
	'data = Replace(data, """", "", 1, -1, 1)
	data = Replace(data, "=", "&#061;", 1, -1, 1)
	'data = Replace(data, "'", "&#146;", 1, -1, 1)

	data = Replace(data, "select", "sel&#101;ct", 1, -1, 1)
	data = Replace(data, "join", "jo&#105;n", 1, -1, 1)
	data = Replace(data, "union", "un&#105;on", 1, -1, 1)
	data = Replace(data, "where", "wh&#101;re", 1, -1, 1)
	data = Replace(data, "insert", "ins&#101;rt", 1, -1, 1)
	data = Replace(data, "delete", "del&#101;te", 1, -1, 1)
	data = Replace(data, "update", "up&#100;ate", 1, -1, 1)
	data = Replace(data, "like", "lik&#101;", 1, -1, 1)
	data = Replace(data, "drop", "dro&#112;", 1, -1, 1)
	data = Replace(data, "create", "cr&#101;ate", 1, -1, 1)
	data = Replace(data, "modify", "mod&#105;fy", 1, -1, 1)
	data = Replace(data, "rename", "ren&#097;me", 1, -1, 1)
	data = Replace(data, "alter", "alt&#101;r", 1, -1, 1)
	data = Replace(data, "cast", "ca&#115;t", 1, -1, 1)
    temizle= data
End Function

Private Function temizle1(ByVal data)
	data = Replace(data, "&#097;", "a", 1, -1, 0)
	data = Replace(data, "&#098;", "b", 1, -1, 0)
	data = Replace(data, "&#099;", "c", 1, -1, 0)
	data = Replace(data, "&#100;", "d", 1, -1, 0)
	data = Replace(data, "&#101;", "e", 1, -1, 0)
	data = Replace(data, "&#102;", "f", 1, -1, 0)
	data = Replace(data, "&#103;", "g", 1, -1, 0)
	data = Replace(data, "&#104;", "h", 1, -1, 0)
	data = Replace(data, "&#105;", "i", 1, -1, 0)
	data = Replace(data, "&#106;", "j", 1, -1, 0)
	data = Replace(data, "&#107;", "k", 1, -1, 0)
	data = Replace(data, "&#108;", "l", 1, -1, 0)
	data = Replace(data, "&#109;", "m", 1, -1, 0)
	data = Replace(data, "&#110;", "n", 1, -1, 0)
	data = Replace(data, "&#111;", "o", 1, -1, 0)
	data = Replace(data, "&#112;", "p", 1, -1, 0)
	data = Replace(data, "&#113;", "q", 1, -1, 0)
	data = Replace(data, "&#114;", "r", 1, -1, 0)
	data = Replace(data, "&#115;", "s", 1, -1, 0)
	data = Replace(data, "&#116;", "t", 1, -1, 0)
	data = Replace(data, "&#117;", "u", 1, -1, 0)
	data = Replace(data, "&#118;", "v", 1, -1, 0)
	data = Replace(data, "&#119;", "w", 1, -1, 0)
	data = Replace(data, "&#120;", "x", 1, -1, 0)
	data = Replace(data, "&#121;", "y", 1, -1, 0)
	data = Replace(data, "&#122;", "z", 1, -1, 0)

	data = Replace(data, "&#065;", "A", 1, -1, 0)
	data = Replace(data, "&#066;", "B", 1, -1, 0)
	data = Replace(data, "&#067;", "C", 1, -1, 0)
	data = Replace(data, "&#068;", "D", 1, -1, 0)
	data = Replace(data, "&#069;", "E", 1, -1, 0)
	data = Replace(data, "&#070;", "F", 1, -1, 0)
	data = Replace(data, "&#071;", "G", 1, -1, 0)
	data = Replace(data, "&#072;", "H", 1, -1, 0)
	data = Replace(data, "&#073;", "I", 1, -1, 0)
	data = Replace(data, "&#074;", "J", 1, -1, 0)
	data = Replace(data, "&#075;", "K", 1, -1, 0)
	data = Replace(data, "&#076;", "L", 1, -1, 0)
	data = Replace(data, "&#077;", "M", 1, -1, 0)
	data = Replace(data, "&#078;", "N", 1, -1, 0)
	data = Replace(data, "&#079;", "O", 1, -1, 0)
	data = Replace(data, "&#080;", "P", 1, -1, 0)
	data = Replace(data, "&#081;", "Q", 1, -1, 0)
	data = Replace(data, "&#082;", "R", 1, -1, 0)
	data = Replace(data, "&#083;", "S", 1, -1, 0)
	data = Replace(data, "&#084;", "T", 1, -1, 0)
	data = Replace(data, "&#085;", "U", 1, -1, 0)
	data = Replace(data, "&#086;", "V", 1, -1, 0)
	data = Replace(data, "&#087;", "W", 1, -1, 0)
	data = Replace(data, "&#088;", "X", 1, -1, 0)
	data = Replace(data, "&#089;", "Y", 1, -1, 0)
	data = Replace(data, "&#090;", "Z", 1, -1, 0)

	data = Replace(data, "&#048;", "0", 1, -1, 0)
	data = Replace(data, "&#049;", "1", 1, -1, 0)
	data = Replace(data, "&#050;", "2", 1, -1, 0)
	data = Replace(data, "&#051;", "3", 1, -1, 0)
	data = Replace(data, "&#052;", "4", 1, -1, 0)
	data = Replace(data, "&#053;", "5", 1, -1, 0)
	data = Replace(data, "&#054;", "6", 1, -1, 0)
	data = Replace(data, "&#055;", "7", 1, -1, 0)
	data = Replace(data, "&#056;", "8", 1, -1, 0)
	data = Replace(data, "&#057;", "9", 1, -1, 0)
	
	data = Replace(data, "&#061;", "=", 1, -1, 0)
	data = Replace(data, "&lt;", "<", 1, -1, 0)
	data = Replace(data, "&gt;", ">", 1, -1, 0)
	data = Replace(data, "&amp;", "&", 1, -1, 0)
	data = Replace(data, "&#146;", "'", 1, -1, 1)

	temizle1 = data
End Function


Public  Function htmltemizle(data)
If data= "" Then Exit Function
data =  Replace(data, "<br>", VBCRLF)
arrKod =  Split(data, "<")
FOR  EACH Kod  IN  arrKod
strGidecek = strGidecek &  Mid(Kod,  Instr(Kod, ">")+1,  Len(Kod)-Instr(Kod, ">"))
NEXT
htmltemizle=strGidecek
End  Function





'Function htmltemizle(data)
  'Set objReg = New RegExp 
  'objReg.Global = True 
  'objReg.IgnoreCase = True 
  'objReg.Pattern = "<[>]+>" 
  'YeniText = objReg.Replace(text,"") 
  'Set objReg = Nothing
'htmltemizle = YeniText
'End Function






function  encode(Text) 
Dim  TextCharCode, PasswordCharCode, NewCharCode
Password="efkan"
For  Char = 1 To  LEN(Text) 
TextCharCode =  ASC(MID(Text,Char,1))
PasswordCharCode =  ASC(MID(Password,(Char MOD  LEN(Password) + 1),1))
NewCharCode = TextCharCode + PasswordCharCode
if  NewCharCode > 255  Then  NewCharCode = NewCharCode -255
encode = encode &  CHR(NewCharCode)
NEXT
End  function

function  decode(Code) 
Dim  CodeCharCode, PasswordCharCode, OriginalCharCode
Password="efkan"
For  Char = 1 To  LEN(Code)
CodeCharCode =  ASC(MID(Code,Char,1))
PasswordCharCode =  ASC(MID(Password,(Char MOD  LEN(Password) + 1),1))
OriginalCharCode = CodeCharCode - PasswordCharCode
if OriginalCharCode < 1  Then  OriginalCharCode = OriginalCharCode + 255
decode = decode &  CHR(OriginalCharCode)
NEXT
End  Function








function kontrol(data)
If data= "" Then Exit Function
data = Trim(data)
IF Not IsNumeric(data) THEN
Response.Write "<B>Lütfen Geçerli bir ID numarasý girin.</B>"
Response.End
End If
kontrol=data
end function



function kodver(data)
session("gkodu")=""
minsayi = 10000 'seçilecek sayýnýn alt sýnýrý
maxsayi = 99999 'seçilecek sayýnýn üst sýnýrý
Randomize()
intRangeSize = maxsayi - minsayi + 1
sngRandomValue = intRangeSize * Rnd()
sngRandomValue = sngRandomValue + minsayi
data = Int(sngRandomValue)
session("gkodu")=trim(data)
kodver = trim(data)
end function

function kodver2(data)
session("gkodu2")=""
minsayi = 10000 'seçilecek sayýnýn alt sýnýrý
maxsayi = 99999 'seçilecek sayýnýn üst sýnýrý
Randomize()
intRangeSize = maxsayi - minsayi + 1
sngRandomValue = intRangeSize * Rnd()
sngRandomValue = sngRandomValue + minsayi
data = Int(sngRandomValue)
session("gkodu2")=trim(data)
kodver2 = trim(data)
end function





Public  Function Vurgula (ByRef sGelen, sKelime)
If sgelen= "" Then Exit Function
sGelen=Replace(sGelen, sKelime, "<span style=""background-color: #FFFF00"">"& sKelime &"</span>")
'sGelen=Replace(sGelen, sKelime, "<b>"& sKelime &"</b>")
Vurgula=sGelen
End  Function


function buyukharf(bharf)
Session.LCID = 1055
Session.CodePage = 1254
DefaultLCID = Session.LCID 
DefaultCodePage = Session.CodePage
If bharf= "" Then Exit Function
bharf = Trim(bharf)
bharf = Replace(bharf , "ý", "I")
bharf = Replace(bharf , "i", "Ý")
bharf = Replace(bharf , "Þ", "Þ")
bharf = Replace(bharf , "Ý", "Ý")
bharf = Replace(bharf , "Ç", "Ç")
bharf = Replace(bharf , "ç", "Ç")
bharf = Ucase(bharf)
buyukharf=bharf
end function




function kucukharf(kharf)
If kharf= "" Then Exit Function
kharf = Trim(kharf)
kharf = Replace(kharf , "I", "ý")
kharf = Replace(kharf , "Ý", "i")
kharf = Lcase(kharf)
basharf=Left(kharf,1)
basharf=Replace(basharf , "ý", "I")
basharf=Replace(basharf , "i", "Ý")
basharf = Ucase(basharf)
kharf =Mid(kharf,2,5000000)
kucukharf=basharf & kharf

end function


Function  linkyap(byVal Text)
If  Text="" OR  isNull(Text)  Then  Exit  Function
'Replace Breaks  to  <br />s
Text =  Replace(Text,Chr(13),"<br /> ")
'//  Set  Array
Dim  LinkArr, i
LinkArr =  Split(Text," ")
'//  Loop  & Replace
For  i=0 to  Ubound(LinkArr) 
If  Instr(Lcase(Left(LinkArr(i),15)),"http://") OR  Instr(Lcase(Left(LinkArr(i),15)),"www.")  Then
LinkArr(i) = "<a href=""" & LinkArr(i) & """>" & LinkArr(i) & "</a>"
End  If
Next
'// Join & Return
linkyap = Join(LinkArr," ")
End  Function



Function emailkontrol(strVeri)
If strVeri = "" Then Exit Function
Set objRegExp = New Regexp
With objRegExp
          .Pattern = "[a-zA-Z0-9]+@[a-zA-Z0-9]+\.[a-zA-Z0-9]+"
          .IgnoreCase = False
          .Global = True
End With
If objRegExp.Test(strVeri) = True Then
emailkontrol= True
Else
emailkontrol = False
End If
End Function



function suz(mesaj)
kelimeler = "salak,aptal,manyak"
hata =  split(kelimeler, ",")
for  i = 0 to  ubound(hata)
mesaj =  Replace(mesaj, hata(i),  string(len(hata(i)),"*"), 1,-1,1) 
next
suz=mesaj
end function

%>



<script language="JavaScript">
<!--
function submitConfirm(listForm)
{ 
   listForm.target="_self"; 
   listForm.action="";
   var answer = confirm ("Bu kayýdý / kayýtlarý silmek istediðinize eminmisiniz?") 
   if (answer)
       return true;
   else
       return false;
} 
//-->
</script>

<SCRIPT type=text/javascript>
<!--
function toggle(theElem){
document.getElementById(theElem).style.display = (document.getElementById(theElem).style.display == 'none')?'':'none';
}
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</SCRIPT>