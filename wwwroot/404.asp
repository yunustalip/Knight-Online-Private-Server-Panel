<!--#include file="function.asp"-->
<%Response.Write "<base href=""http://"&Request.ServerVariables("server_name")&""">"
Response.Expires=0
Response.Charset="iso-8859-9"
Url = Replace((Trim(Request.Querystring)),"404;","")
gelenLink_bol = split(Url, "/")

If UBound(gelenlink_bol)<3 Then
Response.End
End If
Link=lcase(gelenLink_bol(3))

Uzanti=InStr(Link,".")
If Uzanti=0 Then
Uzanti=Len(Link)
Else
Uzanti=Uzanti-1
End If

Session("Sayfa")=Replace(Url,".html","")

Link=Mid(Link,1,Uzanti)

If Link="anasayfa" Then
Server.Execute("Default.Asp")
Response.End
End If

If Link="user-ranking" Then
Sayfa="UserRanking.Asp"
ElseIf Link="monthly-ranking" Then
Sayfa="MonthlyRanking.asp"
ElseIf Link="weekly-ranking" Then
Sayfa="WeeklyRanking.asp"
ElseIf Link="daily-ranking" Then
Sayfa="DailyRanking.asp"
ElseIf Link="ardream-ranking" Then
Sayfa="ArdreamRanking.asp"
ElseIf Link="clan-ranking" Then
Sayfa="ClanRanking.asp"

ElseIf Link="online" Then
Sayfa="online.asp"
ElseIf Link="register" Then
Sayfa="register.asp"

ElseIf Link="clan-np-donate-ranking" Then
Sayfa="ClanNpDonateRanking.asp"
ElseIf Link="clan-detay" Then
Sayfa="/sayfalar/showclan.asp"
ElseIf Link="clan-np-detay" Then
Sayfa="/sayfalar/showclannpdetay.asp"
ElseIf Link="karakter-detay" Then
Sayfa="KarakterDetay.asp"
ElseIf Link="userbilgi" Then
Sayfa="userbilgi.asp"
ElseIf Link="kim-kimi-kesmis" Then
Sayfa="kimkimikesmis.asp"
ElseIf Link="ban-list" Then
Sayfa="banlist.asp"
ElseIf Link="kim-nerede" Then
Sayfa="kim-nerede.Asp"
ElseIf Link="king-election" Then
Sayfa="king_election.asp"
ElseIf Link="statistics" Then
Sayfa="statistics.asp"
ElseIf Link="search" Then
Sayfa="search.asp"
ElseIf Link="drop-list" Then
Sayfa="droplist.asp"
ElseIf Link="news" Then
Sayfa="news.asp"
ElseIf Link="download" Then
Sayfa="download.asp"
ElseIf Link="chat" Then
Sayfa="Chat/Default2.asp"
ElseIf Link="anasayfa" Then
Sayfa="default.asp"

Else
Session("sayfa")="404 Sayfa Bulunamadý: " & gelenLink_bol(3)
Response.Write "<style>.hatasayfa{font-family:Verdana, Arial, Helvetica, sans-serif;"& vbcrlf
Response.Write "font-size:10px;"& vbcrlf
Response.Write "}</style>"& vbcrlf
Response.Write "<div class=""hatasayfa"">Sayfa Bulunamadý</div>"
Response.End
End If

If Right(Url,5)=".html" Then
Server.Execute(Sayfa)
Response.End

Else
Server.Execute("Default.Asp")
Session("izin")=""

For LT = 3 To UBound(gelenlink_bol)
Adresa = Adresa &"/"& gelenlink_bol(Lt)
Next
Response.Write "<script>openpage('"&Adresa&".html')</script>"


End If%>

