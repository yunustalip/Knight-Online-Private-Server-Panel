<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<%Response.expires=0 
Dim MenuAyar,ksira,REFERER_URL,s,REFERER_DOMAIN,link,gelenlink_bol,d1,d2,d3,d4,csira,tp,siralama,clan,clantp,toplamclan,sira,humanclan,karusclan,clanid,name,members,leader,clannation,totalnp,ranking,cape,flag,grade,clangrade,derece
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='ClanRanking'")
If MenuAyar("PSt")=1 Then
If Not Request.ServerVariables("Script_Name")="/404.asp"  Then
yn("/Clan-Ranking")
End If

REFERER_URL = Request.ServerVariables("HTTP_REFERER")
s=Request.ServerVariables("script_name")
if InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
else
REFERER_DOMAIN = left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If 

if REFERER_DOMAIN="http://"&Request.ServerVariables("server_name") or  REFERER_DOMAIN="http://www."&Request.ServerVariables("server_name") Then
else

yn("/Clan-Ranking")
End If

Dim siteayr
Set siteayr=Conne.Execute("select clansiralama from siteayar")
csira=siteayr("clansiralama")

link = Session("Sayfa")
gelenlink_bol = split(link, "/")
tp=ubound(gelenlink_bol)

if tp=4 Then
d1=gelenlink_bol(4)
elseif tp=5 Then
d1=gelenlink_bol(4)
d2=gelenlink_bol(5)
elseif tp=6 Then
d1=gelenlink_bol(4)
d2=gelenlink_bol(5)
d3=gelenlink_bol(6)
elseif tp=7 Then
d1=gelenlink_bol(4)
d2=gelenlink_bol(5)
d3=gelenlink_bol(6)
d4=gelenlink_bol(7)
End If

siralama=secur(d1)
if siralama="" Then
siralama="ortak"
End If

if siralama="karus" Then
Response.Write "<br><center><img src=""imgs/karusclanranking.gif"" alt="""">"
elseif siralama="elmorad" Then
Response.Write "<br><center><img src=""imgs/humanclanranking.gif"" alt="""">"
else
Response.Write "<br><center><img src=""imgs/clanranking.gif"" alt="""">"
End If%><style>
.link1:link{
	color: #808080;
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
	text-decoration: none;
	font-weight: bold;
	
}
.link1:hover {
	color:#FF0000;
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
	text-decoration: none;
	font-weight: bold;
}
.link1:active{
	color: #808080;
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
	text-decoration: none;
	font-weight: bold;
	
}
.link1:visited {
	color: #808080;
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
	text-decoration:none;
	font-weight: bold;
}
.link1:visited:hover {
	color:#FF0000;
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
	text-decoration: none;	
	font-weight: bold;
}

.link1:visited:active {
	color: #808080;
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
	font-weight: bold;
}
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}

td {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
}
</style><br><br>
<br><%


If siralama="ortak" Then

Set clan = Conne.Execute("Select Top "&csira&" k.IDNum, k.IDName, k.Members, k.Chief, k.Nation, SUM(u.Loyalty), k.ranking, k.scape, k.flag From KNIGHTS k,USERDATA u,knights_user ku Where k.IDNum = ku.sIDNum AND ku.strUserID = u.strUserID and u.authority<>0 and u.authority<>255  GROUP BY k.IDNum,k.Nation,k.IDName,k.Members,k.Chief,k.ranking,k.scape,k.flag ORDER BY SUM(u.Loyalty) DESC")
Set clantp = Conne.Execute("Select count(*) as toplamclan From KNIGHTS")
toplamclan=clantp("toplamclan")

ElseIf siralama="elmorad" Then

Set clan = Conne.Execute("Select Top "&csira&" k.IDNum, k.IDName, k.Members, k.Chief, k.Nation, SUM(u.Loyalty), k.ranking,k.scape, k.flag From KNIGHTS k,USERDATA u,knights_user ku Where k.IDNum = ku.sIDNum AND ku.strUserID = u.strUserID and u.authority<>0 and u.authority<>255 and k.nation=2 GROUP BY k.IDNum,k.Nation,k.IDName,k.Members,k.Chief,k.ranking,k.scape,k.flag ORDER BY SUM(u.Loyalty) DESC, k.ranking DESC")
Set clantp = Conne.Execute("Select count(*) toplamclan From KNIGHTS")
toplamclan=clantp("toplamclan")

ElseIf siralama="karus" Then

Set clan = Conne.Execute("Select Top "&csira&" k.IDNum, k.IDName, k.Members, k.Chief, k.Nation, SUM(u.Loyalty), k.ranking, k.scape, k.flag From KNIGHTS k,USERDATA u,knights_user ku Where k.IDNum = ku.sIDNum AND ku.strUserID = u.strUserID and u.authority<>0 and u.authority<>255 and k.nation=1  GROUP BY k.IDNum,k.Nation,k.IDName,k.Members,k.Chief,k.ranking,k.scape,k.flag ORDER BY SUM(u.Loyalty) DESC, k.ranking DESC")
Set clantp = Conne.Execute("Select count(*) toplamclan From KNIGHTS")
toplamclan = clantp("toplamclan")
End If
%>
<b><img src="imgs/ilk5_grade5.gif" width="20" height="20" align="absmiddle">&nbsp;Ilk 5 Clan(Yanan Kolluk)&nbsp;&nbsp;&nbsp;&nbsp;<img src="imgs/ust_grade5.gif" width="20" height="20" align="absmiddle">&nbsp;Üst Clanlar &nbsp;&nbsp;&nbsp;&nbsp;<img src="imgs/alt_grade5.gif" width="20" height="20" align="absmiddle">&nbsp;Alt Clanlar</b><br />
<br><b><a href="/Clan-Ranking" onclick="pageload('Clan-Ranking');return false" class="link1"><img src="imgs/ortak.gif"  border="0" align="texttop">&nbsp;Ortak Sýralama</a>&nbsp;-&nbsp;<a href="/Clan-Ranking/karus" onclick="javascript:pageload('/Clan-Ranking/karus');return false" class="link1"><img src="imgs/karuslogo.gif" border="0" align="texttop">&nbsp;Karus Sýralamasý</a>&nbsp;-&nbsp;<a href="/Clan-Ranking/elmorad" onclick="javascript:pageload('/Clan-Ranking/elmorad');return false" class="link1"><img src="imgs/elmologo.gif" width="15" height="15" border="0" align="absmiddle">&nbsp;Human Sýralamasý</a></b>
<br>
<b>Toplam Clan : </b> <%=toplamclan%><br />
<table width="630" border="0" align="center">
  <tr>
	<td width="30" height="16" align="center" background="imgs/menubg.gif" ><span class="style1">Sýra </span></td>
	<td width="184" align="center" background="imgs/menubg.gif" ><span class="style1">Clan Adý</span></td>	
	<td width="200" align="center" background="imgs/menubg.gif" ><span class="style1">NP</span></td>
	<td width="184" align="center" background="imgs/menubg.gif" ><span class="style1">Clan Lideri</span></td>	
	<td width="50" align="center" background="imgs/menubg.gif" ><span class="style1">Grade</span></td>
	<td width="80" align="center" background="imgs/menubg.gif" ><span class="style1">Üye Sayýsý</span></td>
	<td width="50" align="center" background="imgs/menubg.gif" ><span class="style1">Irk</span></td>
  </tr>
<%if not clan.eof Then
sira=1
humanclan=0
karusclan=0
do while not clan.eof

clanid = clan(0)
name = clan(1)
members = clan(2)
leader = clan(3)
clannation = clan(4)
totalnp = clan(5)
ranking = clan(6)
cape= clan(7)
flag= clan(8)

if clannation="1" Then
karusclan=karusclan+1
elseif clannation="2" Then
humanclan=humanclan+1
End If

if totalnp<72000 Then
grade=5
elseif totalnp<144000 Then
grade=4
elseif totalnp<360000 Then
grade=3
elseif totalnp<720000 Then
grade=2
elseif totalnp>=720000 Then
grade=1
End If

if clannation="1" and karusclan<6 or clannation="2" and humanclan<6 Then
clangrade="ilk"
else
clangrade="diger"
End If

if grade=5 and flag="2" and clangrade="diger" Then
derece="<img src=imgs/ust_grade5.gif height=20 width=20>"
elseif grade=4 and flag="2" and clangrade="diger" Then
derece="<img src=imgs/ust_grade4.gif height=20 width=20>"
elseif grade=3 and flag="2" and clangrade="diger" Then
derece="<img src=imgs/ust_grade3.gif height=20 width=20>"
elseif grade=2 and flag="2" and clangrade="diger" Then
derece="<img src=imgs/ust_grade2.gif height=20 width=20>"
elseif grade=1 and flag="2" and clangrade="diger" Then
derece="<img src=imgs/ust_grade1.gif height=20 width=20>"
End If

if grade=5 and flag="1" and clangrade="diger" Then
derece="<img src=imgs/alt_grade5.gif height=20 width=20>"
elseif grade=4 and flag="1" and clangrade="diger" Then
derece="<img src=imgs/alt_grade4.gif height=20 width=20>"
elseif grade=3 and flag="1" and clangrade="diger" Then
derece="<img src=imgs/alt_grade3.gif height=20 width=20>"
elseif grade=2 and flag="1" and clangrade="diger" Then
derece="<img src=imgs/alt_grade2.gif height=20 width=20>"
elseif grade=1 and flag="1" and clangrade="diger" Then
derece="<img src=imgs/alt_grade1.gif height=20 width=20>"

End If


if grade=5 and clangrade="ilk" Then
derece="<img src=imgs/ilk5_grade5.gif height=20 width=20>"
elseif grade=4 and clangrade="ilk" Then
derece="<img src=imgs/ilk5_grade4.gif height=20 width=20>"
elseif grade=3 and clangrade="ilk" Then
derece="<img src=imgs/ilk5_grade3.gif height=20 width=20>"
elseif grade=2 and clangrade="ilk" Then
derece="<img src=imgs/ilk5_grade2.gif height=20 width=20>"
elseif grade=1 and clangrade="ilk" Then
derece="<img src=imgs/ilk5_grade1.gif height=20 width=20>"
End If

%>
  <tr bgcolor="#F3D78B" onmouseover="this.style.background='#D5AB4A'" onmouseout="this.style.background='#F3D78B'">
	<td align="center"><%Response.Write(sira)%></td>
	<td align="center"><a href="/Clan-Detay/<%=trim(name)&","&clanid%>" class="link1" style="display:block"  onclick="pageload('/Clan-Detay/<%=name&","&clanid%>');return false"><%=name%></a></td>
	<td align="center"><%=ayir(totalnp)%></td>
	<td align="center"><a href="Karakter-Detay/<%=trim(leader)%>" class="link1" style="display:block" onclick="pageload('Karakter-Detay/<%=trim(leader)%>');return false"><%=leader%></a></td>
	<td align="center"><%=derece%></td>
	<td align="center"><%=members%></td>
	<td align="center"><%=nation(clannation)%></td>

  </tr>
  <%

  sira=sira+1
  clan.MoveNext
  Loop
  clan.close
  set clan=nothing
  clantp.close
  set clantp=nothing
  else
   Response.Write("<table><tr><td align='center'>Kurulu Clan Bulunumadý. </td></tr></table>")
  End If %>
  <tr><td colspan="6" align="center">Ilk <%=csira%> Clan Gösteriliyor.<br />
<br />Not: Anlýk Clan sýralamasýdýr.Oyun içi gradeler resetlerde güncellenir.
  <td></tr>
</table>
<%else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%></center>