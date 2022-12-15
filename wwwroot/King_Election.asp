<!--#include file="_inc/Conn.Asp"-->
<!--#include file="Function.Asp"-->
<!--#include file="Guvenlik.Asp"-->
<style>
body,td{
color: #000000;
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:10px;

}
.baslik{
color: #000000;
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:10px;
font-weight:bold
}
a{
color:#808080;
font-weight:bold
}
a:visited{
color:#808080;
font-weight:bold
}
</style><br><img src="imgs/kingelection.gif"><br><br><br>
<table align="left" style="position:relative;left:50px"><tr><td colspan="3" align="center" class="baslik">Human Kraliyet Adaylýðý Sonuçlarý
<%Dim MenuAyar,ksira,kral
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='King_Election'")
If MenuAyar("PSt")=1 Then

Function BinaryToString(Binary)
  Dim I, S
  For I = 1 To LenB(Binary)
    S = S & Chr(AscB(MidB(Binary, I, 1)))
  Next
  BinaryToString = S
End Function

Set kral=Conne.Execute("select * from KING_ELECTION_LIST where bynation=2")
Do While Not kral.Eof
Dim Oylar,toplamoy,topoy,yuzde,kralv,vaad,x,kralvaad,kral2,oylar2,toplamoy2,topoy2,yuzde2
Set Oylar=Conne.Execute("select count(*) as oy from KING_BALLOT_BOX where bynation=2 and strcandidacyid='"&kral("strname")&"'")
Set Toplamoy=Conne.Execute("select count(*) as toplamoy from KING_BALLOT_BOX where bynation=2")
If Not oylar.Eof And Not toplamoy.Eof Then
topoy=toplamoy(0)
If topoy=0 Then
topoy=1
End If
yuzde=round(oylar("oy")/topoy*100)

Set kralv=Conne.Execute("select * from KING_CANDIDACY_NOTICE_BOARD where struserid='"&kral("strname")&"'")
If Not kralv.Eof Then
vaad=split(BinaryToString(kralv("strnotice")),"#")
For x=0 To ubound(vaad)-1
kralvaad=kralvaad&"<li>"&vaad(x)&"</li>"
Next
Else
kralvaad="Bu Aday Herhangi Bir vaadde Bulunmamýþ."
End If

Response.Write "<tr><td width=""110""><a onMouseOver=""return overlib('<font style=color:#FFFFFF>"&Server.HtmlEncode(kralvaad)&"</font>', RIGHT, WIDTH, 240,CELLPAD, 5, 10, 10)"" onMouseOut=""return nd();"" href=""Karakter-Detay/"&kral("strname")&""" onclick=""pageload('Karakter-Detay/"&trim(kral("strname"))&"');nd();return false"">"&kral("strname")&"</a></td><td width=""100""><img src=""imgs/yuzde.gif"" height=""10"" width="""&yuzde&"""></td><td>%"&yuzde&" ("&oylar("oy")&")</td></tr>"&vbcrlf
End If
Kral.MoveNext
Loop
%>
</table>
<table align="right" style="position:relative;right:50"><tr><td colspan="3" align="center" class="baslik">Karus Kraliyet Adaylýðý Sonuçlarý
<%Set Kral2=Conne.Execute("select * from KING_ELECTION_LIST where bynation=1")
Do While Not Kral2.Eof

Set oylar2=Conne.Execute("select count(*) as oy from KING_BALLOT_BOX where bynation=1 and strcandidacyid='"&kral2("strname")&"'")
set toplamoy2=Conne.Execute("select count(*) as toplamoy from KING_BALLOT_BOX where bynation=1")
if not oylar2.eof and not toplamoy2.eof Then
topoy2=toplamoy2(0)
if topoy2=0 Then
topoy2=1
End If
yuzde2=round(oylar2("oy")/topoy2*100)
set kralv=Conne.Execute("select * from KING_CANDIDACY_NOTICE_BOARD where struserid='"&kral2("strname")&"'")
if not kralv.eof Then
vaad=split(BinaryToString(kralv("strnotice")),"#")
for x=0 to ubound(vaad)
if x=10 Then
Response.Write vaad(x)
else
kralvaad="<li>"&vaad(x)
End If
next
else
kralvaad="Bu Aday Herhangi Bir vaadde Bulunmamýþ."
End If

Response.Write "<tr><td width=""110""><a onMouseOver=""return overlib('<font style=color:#FFFFFF>"&kralvaad&"</font>', RIGHT, WIDTH, 240,CELLPAD, 5, 10, 10)"" onMouseOut=""return nd();"" href=""Karakter-Detay/"&kral2("strname")&""" onclick=""pageload('Karakter-Detay/"&Trim(kral2("strname"))&"');nd();return false"">"&kral2("strname")&"</a></td><td width=""100""><img src=""imgs/yuzde.gif"" height=""10"" width="""&yuzde2&"""></td><td>%"&yuzde2&" ("&oylar2("oy")&")</td></tr>"&vbcrlf
End If
kral2.MoveNext
Loop


MenuAyar.Close
Set MenuAyar=Nothing

Else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
%>
</table>