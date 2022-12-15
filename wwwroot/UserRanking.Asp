<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<%Response.Buffer = True
Dim MenuAyar,ksira,REFERER_URL,s,REFERER_DOMAIN,link,gelenlink_bol,tp,d1,d2,d3,d4,siralama,siralamatur,listeleme,tur,userrank,userranktp,toplamchar,humansira,karussira,sira,clan,clanid,idnum,struserid,style,csira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='UserRanking'")
If MenuAyar("PSt")=1 Then
If Not Request.ServerVariables("Script_Name")="/404.asp" Then
yn("/User-Ranking")
End If

REFERER_URL = Request.ServerVariables("HTTP_REFERER")

If InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
Else
REFERER_DOMAIN = Left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If


If REFERER_DOMAIN="http://"&Request.ServerVariables("server_name") or  REFERER_DOMAIN="http://www."&Request.ServerVariables("server_name") Then
Else
yn("/User-Ranking")
End If

If Instr(Request.ServerVariables("ALL_HTTP"),"HTTP_X_REQUESTED_WITH:")>0 Then
Else
yn("/User-Ranking")
End If

link = Session("Sayfa")
gelenlink_bol = split(link, "/")
tp=ubound(gelenlink_bol)

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

Dim siteayr
Set siteayr=Conne.Execute("select kulsiralama from siteayar")
ksira=siteayr("kulsiralama")

If siralama="" Then
siralama="ortak"
ElseIf siralama<>"karus" and siralama<>"elmorad" Then
siralama="ortak"
End If

If siralamatur="Warrior" and siralama="karus" Then
tur="class=101 and authority=1 or class=105 and authority=1 or class=106 and authority=1"
ElseIf siralamatur="Warrior" and siralama="elmorad" Then
tur="class=201 and authority=1 or class=205 and authority=1 or class=206 and authority=1"
ElseIf siralamatur="Warrior" and siralama="ortak" Then
tur="class=101 and authority=1 or class=105 and authority=1 or class=106 and authority=1 or class=201 and authority=1 or class=205 and authority=1 or class=206 and authority=1"

ElseIf siralamatur="Rogue" and siralama="karus" Then
tur="class=102 and authority=1 or class=107 and authority=1 or class=108 and authority=1" 
ElseIf siralamatur="Rogue" and siralama="elmorad" Then
tur="class=202 and authority=1 or class=207 and authority=1 or class=208 and authority=1"
ElseIf siralamatur="Rogue" and siralama="ortak" Then
tur="class=102 and authority=1 or class=107 and authority=1 or class=108 and authority=1 or class=202 and authority=1 or class=207 and authority=1 or class=208 and authority=1"

ElseIf siralamatur="Priest" and siralama="karus" Then
tur="class=104 and authority=1 or class=111 and authority=1 or class=112 and authority=1"
ElseIf siralamatur="Priest" and siralama="elmorad" Then
tur="class=204 and authority=1 or class=211 and authority=1 or class=212 and authority=1"
ElseIf siralamatur="Priest" and siralama="ortak" Then
tur="class=104 and authority=1 or class=111 and authority=1 or class=112 and authority=1 or class=204 and authority=1 or class=211 and authority=1 or class=212 and authority=1"

ElseIf siralamatur="Mage" and siralama="karus" Then
tur="class=103 and authority=1 or class=109 and authority=1 or class=110 and authority=1"
ElseIf siralamatur="Mage" and siralama="elmorad" Then
tur="class=203 and authority=1 or class=209 and authority=1 or class=210 and authority=1"
ElseIf siralamatur="Mage" and siralama="ortak" Then
tur="class=103 and authority=1 or class=109 and authority=1 or class=110 and authority=1 or class=203 and authority=1 or class=209 and authority=1 or class=210 and authority=1"

ElseIf siralama="karus" and siralamatur="" Then
tur="authority=1 and nation=1"
ElseIf siralama="elmorad" and siralamatur="" Then
tur="authority=1 and nation=2"
ElseIf siralama="ortak" and siralamatur="" Then
tur="authority=1"
Else
tur="authority=1"
End If

function smge(ssira)
if ssira=1 Then
smge="<img src=""imgs/1.gif"">"
ElseIf ssira>1 and ssira<5 Then
smge="<img src=""imgs/2.gif"">"
ElseIf ssira>4 and ssira<10 Then
smge="<img src=""imgs/3.gif"">"
ElseIf ssira>9 and ssira<26 Then
smge="<img src=""imgs/4.gif"">"
ElseIf ssira>25 and ssira<51 Then
smge="<img src=""imgs/5.gif"">"
ElseIf ssira>50 and ssira<101 Then
smge="<img src=""imgs/6.gif"">"
End If
end function


if siralama="karus" Then
set userrank=Conne.Execute("Select Top "&ksira&" StrUserId,Level,Nation,Class,Loyalty,Knights From USERDATA where "&tur&" order by loyalty desc, level desc,struserid asc")
Set userranktp=Conne.Execute("Select count(struserid) toplam From USERDATA where "&tur&"")

toplamchar=userranktp("toplam")

ElseIf siralama="elmorad" Then
set userrank=Conne.Execute("Select Top "&ksira&" StrUserId,Level,Nation,Class,Loyalty,Knights From USERDATA where  "&tur&" order by loyalty desc,Level desc,struserid asc")
Set userranktp=Conne.Execute("Select count(struserid) toplam From USERDATA where "&tur&"")

toplamchar=userranktp("toplam")
else 
set userrank=Conne.Execute("Select top "&ksira&" StrUserId,Level,Nation,Class,Loyalty,Knights From USERDATA where  "&tur&" order by loyalty desc,Level DESC,struserid asc")
Set userranktp=Conne.Execute("Select count(struserid) toplam From USERDATA where "&tur&"")

toplamchar=userranktp("toplam")

End If

if siralama="karus" Then
Response.Write "<center><br><img src=""imgs/karusnp.gif"" alt="""">"
ElseIf siralama="elmorad" Then
Response.Write "<center><br><img src=""imgs/humannp.gif"" alt="""">"
else
Response.Write "<center><br><img src=""imgs/ortaknp.gif"" alt="""">"
End If%>
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
</style>
<br><br><br><br>
<b><a href="/User-Ranking" onclick="pageload('/User-Ranking'); return false;" class="link1"><img src="imgs/ortak.gif"  border="0" align="texttop">&nbsp;Ortak Sýralama</a>&nbsp;-
<a href="/User-Ranking/karus" onclick="pageload('/User-Ranking/karus');return false" class="link1"><img src="imgs/karuslogo.gif" border="0" align="texttop">&nbsp;Karus Sýralamasý</a>&nbsp;-
<a href="/User-Ranking/elmorad" onclick="pageload('/User-Ranking/elmorad');return false" class="link1"><img src="imgs/elmologo.gif" border="0" align="absmiddle">&nbsp;Human Sýralamasý</a></b>
<br>
<b>Toplam Karakter : </b> <%=toplamchar%><br />
<table width="569" border="0">
  <tr>
    <td align="center" background="imgs/menubg.gif" class="nosort">Sýra</td>
<%If siralamatur="" Then
Response.Write("<td  align=""center"" background=""imgs/menubg.gif"" class=""nosort"">Simge</td>")
End If%>
    <td width="135" align="center" background="imgs/menubg.gif" class="nosort">Karakter Adý </td>
    <td width="39" align="center" background="imgs/menubg.gif" class="nosort">Level</td>
    <td width="39" align="center" background="imgs/menubg.gif" class="nosort">Irk</td>
    <td width="90" align="center" background="imgs/menubg.gif" class="nosort">Tür</td>
    <td width="140" align="center" background="imgs/menubg.gif" class="nosort">Clan</td>
    <td width="52" align="center" background="imgs/menubg.gif" class="nosort">NP</td>
  </tr>
<% if not userrank.eof Then
humansira=1
karussira=1
sira=1
do while not userrank.eof
if not userrank("Knights")="0" Then
Set clan=Conne.Execute("Select IDNum,IDName From KNIGHTS where IDNum="&userrank("Knights")&"")
if not clan.eof Then
clanid=clan("idname")
idnum=clan("idnum")
else
clanid=""
idnum=""
End If
else
clanid=""
idnum=""
End If
struserid=trim(userrank("struserid"))
%>
<tr bgcolor="#F3D78B" onmouseover="this.style.background='#D5AB4A'" onmouseout="this.style.background='#F3D78B'">
<td align="center"><%=sira%></td>
<% if siralamatur="" Then
Response.Write "<td align=""center"">"
if userrank("Nation")="1" Then
Response.Write smge(karussira)
ElseIf userrank("Nation")="2" Then
Response.Write smge(humansira)
End If 
Response.Write "</td>"
End If%>
    <td align="center"><a href="Karakter-Detay/<%=struserid%>" style="display:block" onclick="pageload('Karakter-Detay/<%=struserid%>');return false" >
      <%if sirarenk(sira)<>"" Then%><span style="<%Response.Write sirarenk(sira)%>">
      <%End If
Response.Write struserid%></span></a></td>
    <td align="center"><%=userrank("Level")%></td>
    <td align="center"><%Response.Write "<a href=""/User-Ranking/"
	if userrank("Nation")=1 Then
	Response.Write "karus"" onclick=""pageload('/User-Ranking/karus');return false""><img src=""imgs/karuslogo.gif"" border=""0""></a>"
	ElseIf userrank("Nation")=2 Then
	Response.Write "elmorad"" onclick=""pageload('/User-Ranking/elmorad');return false""><img src=""imgs/elmologo.gif"" border=""0""></a>"
	Else
	Response.Write(""">")
	End If%></td>
    <td align="center"><a href="/User-Ranking/<%Response.Write siralama&"/"
Response.Write cla(userrank("Class")) %>" onclick="pageload('/User-Ranking/<%Response.Write siralama&"/"
Response.Write cla(userrank("Class")) %>');return false" class="link1"><%=cla(userrank("Class"))%></a></td>
    <td align="center"><a href="#" style="display:block" class="link1" onclick="pageload('Clan-Detay/,<%=idnum%>');return false;"><%=clanid%></a></td>
    <td align="center"><%=ayir(userrank("Loyalty"))%></td>
  </tr>
<%If userrank("nation")="1" Then
karussira=karussira+1
ElseIf userrank("nation")="2" Then
humansira=humansira+1
End If
sira=sira+1
userrank.MoveNext
Loop
userrank.Close
Set userrank=nothing
userranktp.Close
Set userranktp=nothing
clan.Close
Set clan=nothing
  %></table>
<table>
<tr><td colspan="8" align="center">Ilk <%=ksira%> karakter gösteriliyor.<br />
  <br />
  Not: Anlýk Np sýralamasýdýr. Oyun içi semboller resetlerde güncellenir. </td>
</tr>
  </table>
<%Else Response.Write("<table><tr><td>Kayýtlý Kullanýcý Bulunmamaktadýr.</td></tr></table>")
End If %>
    </center>
<%
Else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If

MenuAyar.Close
Set MenuAyar=Nothing%>