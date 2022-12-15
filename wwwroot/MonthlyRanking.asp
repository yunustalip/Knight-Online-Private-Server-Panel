<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<%Response.expires=0
Dim MenuAyar,ksira,REFERER_URL,s,REFERER_DOMAIN,link,gelenlink_bol,tp,d1,d2,d3,d4,siralama,siralamatur,listeleme,siteayr,tur,userrank,userranktp,toplamchar,humansira,karussira,sira,clanid,idnum,style,clan,csira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='MonthlyRanking'")
If MenuAyar("PSt")=1 Then
If Not Request.ServerVariables("Script_Name")="/404.asp" Then
yn("/Monthly-Ranking")
End If

REFERER_URL = Request.ServerVariables("HTTP_REFERER")
s=Request.ServerVariables("Script_Name")
If InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
Else
REFERER_DOMAIN = Left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If


If REFERER_DOMAIN="http://"&Request.ServerVariables("server_name") or  REFERER_DOMAIN="http://www."&Request.ServerVariables("server_name") Then
Else
yn("/Monthly-Ranking")
End If

link = Session("sayfa")
gelenlink_bol = Split(link, "/")
tp=UBound(gelenlink_bol)

If tp=4 Then
d1=gelenlink_bol(4)
ElseIf tp=5 Then
d1=gelenlink_bol(4)
d2=gelenlink_bol(5)
ElseIf tp=6 Then
d1=gelenlink_bol(4)
d2=gelenlink_bol(5)
d3=gelenlink_bol(6)
ElseIf tp=7 Then
d1=gelenlink_bol(4)
d2=gelenlink_bol(5)
d3=gelenlink_bol(6)
d4=gelenlink_bol(7)
End If

siralama = secur(d1)
siralamatur = secur(d2)
listeleme=secur(d3)

Set siteayr=Conne.Execute("select kulsiralama from siteayar")
ksira=siteayr("kulsiralama")

If siralama="" Then
siralama="ortak"
ElseIf siralama<>"karus" And siralama<>"elmorad" Then
siralama="ortak"
End If

If siralamatur="Warrior" And siralama="karus" Then
tur="class=101 And authority=1 or class=105 And authority=1 or class=106 And authority=1"
ElseIf siralamatur="Warrior" And siralama="elmorad" Then
tur="class=201 And authority=1 or class=205 And authority=1 or class=206 And authority=1"
ElseIf siralamatur="Warrior" And siralama="ortak" Then
tur="class=101 And authority=1 or class=105 And authority=1 or class=106 And authority=1 or class=201 And authority=1 or class=205 And authority=1 or class=206 And authority=1"

ElseIf siralamatur="Rogue" And siralama="karus" Then
tur="class=102 And authority=1 or class=107 And authority=1 or class=108 And authority=1" 
ElseIf siralamatur="Rogue" And siralama="elmorad" Then
tur="class=202 And authority=1 or class=207 And authority=1 or class=208 And authority=1"
ElseIf siralamatur="Rogue" And siralama="ortak" Then
tur="class=102 And authority=1 or class=107 And authority=1 or class=108 And authority=1 or class=202 And authority=1 or class=207 And authority=1 or class=208 And authority=1"

ElseIf siralamatur="Priest" And siralama="karus" Then
tur="class=104 And authority=1 or class=111 And authority=1 or class=112 And authority=1"
ElseIf siralamatur="Priest" And siralama="elmorad" Then
tur="class=204 And authority=1 or class=211 And authority=1 or class=212 And authority=1"
ElseIf siralamatur="Priest" And siralama="ortak" Then
tur="class=104 And authority=1 or class=111 And authority=1 or class=112 And authority=1 or class=204 And authority=1 or class=211 And authority=1 or class=212 And authority=1"

ElseIf siralamatur="Mage" And siralama="karus" Then
tur="class=103 And authority=1 or class=109 And authority=1 or class=110 And authority=1"
ElseIf siralamatur="Mage" And siralama="elmorad" Then
tur="class=203 And authority=1 or class=209 And authority=1 or class=210 And authority=1"
ElseIf siralamatur="Mage" And siralama="ortak" Then
tur="class=103 And authority=1 or class=109 And authority=1 or class=110 And authority=1 or class=203 And authority=1 or class=209 And authority=1 or class=210 And authority=1"

ElseIf siralama="karus" And siralamatur="" Then
tur="authority=1 And nation=1"
ElseIf siralama="elmorad" And siralamatur="" Then
tur="authority=1 And nation=2"
ElseIf siralama="ortak" And siralamatur="" Then
tur="authority=1"
Else
tur="authority=1"
End If

Function smge(ssira)
If ssira=1 Then
smge="<img src=imgs/001.gif>"

ElseIf ssira>1 And ssira<5 Then
smge="<img src=imgs/002.gif>"

ElseIf ssira>4 And ssira<10 Then
smge="<img src=imgs/003.gif>"

ElseIf ssira>9 And ssira<26 Then
smge="<img src=imgs/004.gif>"

ElseIf ssira>25 And ssira<51 Then
smge="<img src=imgs/005.gif>"

ElseIf ssira>50 And ssira<101 Then
smge="<img src=imgs/006.gif>"
End If
End Function


If siralama="ortak" Then
Set userrank =conne.Execute("Select Top "&ksira&" Race,StrUserId,Level,Nation,Class,LoyaltyMonthly,Knights From USERDATA where "&tur&" order By LoyaltyMonthly DESC,Level DESC")
Set userranktp = conne.Execute("Select count(*) toplam From USERDATA where "&tur&" ")
toplamchar=userranktp("toplam")

ElseIf siralama="karus" Then
Set userrank =conne.Execute("Select Top "&ksira&" Race,StrUserId,Level,Nation,Class,LoyaltyMonthly,Knights From USERDATA where "&tur&" ORDer By LoyaltyMonthly DESC,Level DESC")
Set userranktp  =conne.Execute("Select count(*) toplam From USERDATA where "&tur&" ")
toplamchar=userranktp("toplam")

ElseIf siralama="elmorad" Then
Set userrank =conne.Execute("Select Top "&ksira&" Race,StrUserId,Level,Nation,Class,LoyaltyMonthly,Knights From USERDATA where "&tur&" ORDer By LoyaltyMonthly DESC,Level DESC")
Set userranktp =conne.Execute("Select count(*) toplam From USERDATA where "&tur&"")
toplamchar=userranktp("toplam")

Else
yn("/Monthly-Ranking")
End If

If siralama="karus" Then
Response.Write "<br><img src=""imgs/karusmonthlynp.gif"" alt="""">"
ElseIf siralama="elmorad" Then
Response.Write "<br><img src=""imgs/humanmonthlynp.gif"" alt="""">"
Else
Response.Write "<br><img src=""imgs/monthlynp.gif"" alt="""">"
End If%><center>
<style type="text/css">
.nosort {color: #FFFFFF; font-weight:bold}
td {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
}
a{
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
	text-decoration:none;
}
</style><br><br><br>
<center><b>
<a href="/Monthly-Ranking" onclick="javascript:pageload('/Monthly-Ranking/ortak');return false" class="link1"><img src="imgs/ortak.gif" border="0" align="texttop">&nbsp;Ortak Sýralama</a>&nbsp;-
<a href="/Monthly-Ranking/karus" onclick="javascript:pageload('/Monthly-Ranking/karus');return false" class="link1"><img src="imgs/karuslogo.gif" border="0" align="texttop">&nbsp;Karus Sýralamasý</a>&nbsp;-
<a href="/Monthly-Ranking/elmorad" onclick="javascript:pageload('/Monthly-Ranking/elmorad');return false" class="link1"><img src="imgs/elmologo.gif" border="0" align="absmiddle">&nbsp;Human Sýralamasý</a></b>
</center>
<b>Toplam Karakter : </b> <%=toplamchar%><br />
<table width="569" border="0"  class="sortable" id="sorter">
  <tr>
    <td align="center" background="imgs/menubg.gif" class="nosort">Sýra</td>
<%If siralamatur="" Then%>
    <td  align="center" background="imgs/menubg.gif" class="nosort">Simge</td><%End If%>
    <td width="135" align="center" background="imgs/menubg.gif" class="nosort">Karakter Adý </td>
    <td width="39" align="center" background="imgs/menubg.gif" class="nosort">Level</td>
    <td width="39" align="center" background="imgs/menubg.gif" class="nosort">Irk</td>
    <td width="90" align="center" background="imgs/menubg.gif" class="nosort">Tür</td>
    <td width="140" align="center" background="imgs/menubg.gif" class="nosort">Clan</td>
    <td width="52" align="center" background="imgs/menubg.gif" class="nosort">NP</td>
  </tr>
<% If not userrank.eof Then
humansira=1
karussira=1
sira=1
do while not userrank.eof 
If not userrank("Knights")="0" Then
Set clan=conne.Execute("Select IDNum,IDName From KNIGHTS where IDNum="&userrank("Knights")&"")
If not clan.eof Then
clanid=clan("idname")
idnum=clan("idnum")
Else
clanid=""
idnum=""
End If
Else
clanid=""
idnum=""
End If
%>
<tr bgcolor="#F3D78B" onmouseover="this.style.background='#D5AB4A'" onmouseout="this.style.background='#F3D78B'">
<td align="center"><%=sira%></td>
<% If siralamatur="" Then
Response.Write "<td align=""center"">"
If userrank("Nation")="1" Then
Response.Write smge(karussira)
ElseIf userrank("Nation")="2" Then
Response.Write smge(humansira)
End If
Response.Write "</td>"
End If%>
    <td align="center"><a href="Karakter-Detay/<%=trim(userrank("strUserId"))%>" style="display:block" onclick="pageload('Karakter-Detay/<%=trim(userrank("strUserId"))%>');return false" ><span style="<%=sirarenk(sira)%>"><%=trim(userrank("strUserId"))%></span></a></td>
    <td align="center"><%=userrank("Level")%></td>
    <td align="center"><% nation(userrank("Nation"))%></td>
    <td align="center"><a href="/Monthly-Ranking/<%Response.Write siralama&"/"
Response.Write cla(userrank("Class")) %>" onclick="javascript:pageload('/Monthly-Ranking/<%Response.Write siralama&"/"
Response.Write cla(userrank("Class")) %>');return false" class="link1"><%=cla(userrank("Class"))%></a></td>
    <td align="center"><a href="#" style="display:block" class="link1" onclick="pageload('sayfalar/showclan.asp?goster=<%=idnum%>');return false;"><%=clanid%></a></td>
    <td align="center"><%=ayir(userrank("LoyaltyMonthly"))%></td>
  </tr>
<%If userrank("nation")="1" Then
karussira=karussira+1
ElseIf userrank("nation")="2" Then
humansira=humansira+1
End If
sira=sira+1
userrank.MoveNext
  Loop
userrank.close
set userrank=nothing
userranktp.close
set userranktp=nothing
clan.close
set clan=nothing
  %></table>
<table>
<tr><td colspan="8" align="center">Ilk <%=ksira%> karakter gösteriliyor.<br />
    <br />
    Not: Anlýk Np sýralamasýdýr. Oyun için semboller resetlerde güncellenir. </td>
  </tr>
  </table>
  <%Else
 Response.Write("<table><tr><td>Karakter Bulunamadý.</td></tr></table>")
 End If %>
	</center>
<%
Else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If 

MenuAyar.Close
Set MenuAyar=Nothing
%>